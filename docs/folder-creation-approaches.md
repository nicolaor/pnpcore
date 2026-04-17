# Folder Creation for Non-Document-Library Lists — Proposed Approaches

## Problem

When creating folders inside a SharePoint **generic list** (or any non-document-library list), the SDK currently always uses the `Folders/AddUsingPath` REST endpoint. This creates a filesystem-level folder but does **not** create the associated list-item metadata, making the folder invisible in the SharePoint list UI.

The correct endpoint for non-document-library lists is `AddSubFolderUsingPath`, which creates both the folder and the list-item metadata. `Folders/AddUsingPath` should only be used for document libraries and web-level folders.

### Current behavior (buggy)

```
POST _api/web/lists('{listId}')/rootfolder/Folders/AddUsingPath(decodedurl='MyFolder')
```

This always succeeds (returns 200) but the folder is **not visible** in the SharePoint list UI for generic lists.

### Expected behavior

For generic lists:
```
POST _api/web/lists('{listId}')/rootfolder/AddSubFolderUsingPath(DecodedUrl='MyFolder')
```

For document libraries (unchanged):
```
POST _api/web/lists('{listId}')/rootfolder/Folders/AddUsingPath(decodedurl='MyFolder')
```

---

Below are three proposed approaches to fix this. All preserve backward compatibility — existing callers continue to work without changes.

---

## Approach A — Auto-detect list type via parent model hierarchy ⭐ Recommended

Modify the `AddApiCallHandler` in the `Folder()` constructor to automatically walk up the in-memory parent model hierarchy, find the owning `List`, check its `BaseType`, and choose the correct endpoint. No public API changes are needed — the fix is entirely internal.

### How it works

1. **`FindParentList` helper** — An iterative walk up the `Parent` pointer chain from the current `Folder` model object. If a `List` ancestor is found, return it; otherwise return `null` (indicating a web-level folder, which uses the existing endpoint).
2. **`EnsurePropertiesAsync(BaseType)`** — Once the parent `List` is found, ensure `BaseType` is loaded (one server call, cached per `List` instance for the session).
3. **Endpoint selection** — If `BaseType != DocumentLibrary`, use `AddSubFolderUsingPath`; otherwise use `Folders/AddUsingPath`.

### Files changed

| File | Change |
|------|--------|
| `Folder.cs` (constructor) | Modify `AddApiCallHandler` to call `FindParentList`, check `BaseType`, choose endpoint |
| `Folder.cs` (new helper) | Add `FindParentList(IDataModelParent)` — iterative parent walk |

### Code sketch — `Folder.cs` constructor `AddApiCallHandler`

```csharp
AddApiCallHandler = async (keyValuePairs) =>
{
    var entity = EntityManager.GetClassInfo(GetType(), this);
    string encodedPath = WebUtility.UrlEncode(Name.Replace("'", "''").Replace("%20", " ")).Replace("+", "%20");

    // Walk the parent hierarchy to find the owning List
    var parentList = FindParentList(this);
    if (parentList != null)
    {
        await parentList.EnsurePropertiesAsync(p => p.BaseType).ConfigureAwait(false);

        if (parentList.BaseType != ListBaseType.DocumentLibrary)
        {
            // Generic list, tasks list, etc. — use AddSubFolderUsingPath
            return new ApiCall(
                $"{entity.SharePointGet}/AddSubFolderUsingPath(DecodedUrl='{encodedPath}')",
                ApiType.SPORest);
        }
    }

    // Document library or web-level folder — use existing endpoint
    return new ApiCall(
        $"{entity.SharePointGet}/Folders/AddUsingPath(decodedurl='{encodedPath}')",
        ApiType.SPORest);
};
```

### Code sketch — `FindParentList` helper

```csharp
private static List FindParentList(IDataModelParent model)
{
    var current = model?.Parent;
    while (current != null)
    {
        if (current is List list)
            return list;
        current = (current as IDataModelParent)?.Parent;
    }
    return null;
}
```

### Caller usage — unchanged

```csharp
// Works for document libraries (existing behavior):
await docLib.RootFolder.Folders.AddAsync("MyFolder");

// Now also works correctly for generic lists:
await genericList.RootFolder.Folders.AddAsync("VisibleFolder");

// EnsureFolderAsync also works:
await genericList.RootFolder.EnsureFolderAsync("sub1/sub2");
```

### Pros & Cons

✅ **Zero public API changes** — completely transparent to callers.
✅ **Minimal code change** — only touches `Folder.cs`.
✅ `FindParentList` is an in-memory walk (zero server calls); `EnsurePropertiesAsync(BaseType)` is cached per `List` instance, so creating N folders incurs at most 1 extra round-trip.
❌ Caller cannot bypass the parent walk + `EnsureProperties` code path (though the cost is negligible).

---

## Approach B — New overloads with an optional `ListBaseType?` parameter

Add new overloads alongside the existing ones that accept an explicit `ListBaseType?`. When provided, the SDK skips auto-detection and uses the supplied type directly. When `null` or omitted, falls back to auto-detection (Approach A logic).

### Files changed

| File | Change |
|------|--------|
| `IFolderCollection.cs` | Add 4 new overloads with `ListBaseType? listBaseType = null` |
| `FolderCollection.cs` | Implement the new overloads, pass `listBaseType` through `keyValuePairs` dictionary |
| `IFolder.cs` | Add 4 new overloads for `AddFolderAsync`/`AddFolderBatch*` with `ListBaseType?` |
| `Folder.cs` (extension methods) | Implement the new overloads, forward to `FolderCollection` |
| `Folder.cs` (constructor) | Read `listBaseType` from `keyValuePairs` when present; fall back to auto-detection when `null` |

### Code sketch — `IFolderCollection.cs`

```csharp
// NEW — existing Add(string name) overloads remain unchanged
Task<IFolder> AddAsync(string name, ListBaseType? listBaseType);
IFolder Add(string name, ListBaseType? listBaseType);
Task<IFolder> AddBatchAsync(Batch batch, string name, ListBaseType? listBaseType);
IFolder AddBatch(Batch batch, string name, ListBaseType? listBaseType);
```

### Code sketch — `FolderCollection.cs`

```csharp
public async Task<IFolder> AddAsync(string name, ListBaseType? listBaseType)
{
    if (string.IsNullOrEmpty(name))
        throw new ArgumentNullException(nameof(name));

    var newFolder = CreateNewAndAdd() as Folder;
    newFolder.Name = name;

    // Pass the hint through keyValuePairs (same pattern as ListItemCollection)
    var kvp = listBaseType.HasValue
        ? new Dictionary<string, object> { { "ListBaseType", listBaseType.Value } }
        : null;

    return await newFolder.AddAsync(kvp).ConfigureAwait(false) as Folder;
}
```

### Code sketch — `Folder.cs` constructor `AddApiCallHandler`

```csharp
AddApiCallHandler = async (keyValuePairs) =>
{
    var entity = EntityManager.GetClassInfo(GetType(), this);
    string encodedPath = WebUtility.UrlEncode(Name.Replace("'", "''").Replace("%20", " ")).Replace("+", "%20");

    // 1. Check caller-supplied hint first
    ListBaseType? hintedBaseType = null;
    if (keyValuePairs != null
        && keyValuePairs.TryGetValue("ListBaseType", out var val)
        && val is ListBaseType bt)
    {
        hintedBaseType = bt;
    }

    // 2. Fall back to auto-detection via parent walk
    if (!hintedBaseType.HasValue)
    {
        var parentList = FindParentList(this);
        if (parentList != null)
        {
            await parentList.EnsurePropertiesAsync(p => p.BaseType).ConfigureAwait(false);
            hintedBaseType = parentList.BaseType;
        }
    }

    // 3. Choose endpoint
    if (hintedBaseType.HasValue && hintedBaseType.Value != ListBaseType.DocumentLibrary)
    {
        return new ApiCall($"{entity.SharePointGet}/AddSubFolderUsingPath(DecodedUrl='{encodedPath}')", ApiType.SPORest);
    }

    return new ApiCall($"{entity.SharePointGet}/Folders/AddUsingPath(decodedurl='{encodedPath}')", ApiType.SPORest);
};
```

### Caller usage

```csharp
// Existing code — unchanged, auto-detects:
await list.RootFolder.Folders.AddAsync("MyFolder");

// New code — caller provides the type, skips FindParentList + EnsureProperties:
await list.RootFolder.Folders.AddAsync("MyFolder", ListBaseType.GenericList);
```

### Pros & Cons

✅ Simple, discoverable, follows existing `ListItemCollection` pattern.
✅ Gives callers an explicit "fast path" when they already know the list type.
❌ Adds ~8 overloads across `IFolderCollection`/`IFolder`/`FolderCollection`/`Folder`.

---

## Approach C — `FolderAddOptions` class (more extensible)

Introduce an options object (similar to `MoveCopyOptions` used by `CopyToAsync`/`MoveToAsync`) and add new overloads that accept it. Internally, falls back to auto-detection (Approach A logic) when the option is not set.

### Files changed

| File | Change |
|------|--------|
| New `FolderAddOptions.cs` | New class with `ListBaseType?` property |
| `IFolderCollection.cs` | Add 4 new overloads with `FolderAddOptions options = null` |
| `FolderCollection.cs` | Implement new overloads, pack options into `keyValuePairs` |
| `IFolder.cs` | Add 4 new `AddFolderAsync`/`AddFolderBatch*` overloads |
| `Folder.cs` (extension methods + constructor) | Same as Approach B for the handler logic |

### Code sketch — `FolderAddOptions.cs`

```csharp
public class FolderAddOptions
{
    /// <summary>
    /// Hint for the base type of the parent list. When provided, skips automatic detection
    /// via the parent model hierarchy. When null, the SDK auto-detects.
    /// </summary>
    public ListBaseType? ParentListBaseType { get; set; }
}
```

### Code sketch — `IFolderCollection.cs`

```csharp
Task<IFolder> AddAsync(string name, FolderAddOptions options);
IFolder Add(string name, FolderAddOptions options);
Task<IFolder> AddBatchAsync(Batch batch, string name, FolderAddOptions options);
IFolder AddBatch(Batch batch, string name, FolderAddOptions options);
```

### Caller usage

```csharp
// Existing code — unchanged:
await list.RootFolder.Folders.AddAsync("MyFolder");

// New code with options:
await list.RootFolder.Folders.AddAsync("MyFolder", new FolderAddOptions
{
    ParentListBaseType = ListBaseType.GenericList
});
```

### Pros & Cons

✅ Extensible if more folder-creation options are needed in the future; clean API.
❌ Adds a new public class; still adds ~8 overloads.

---

## Recommendation

**Approach A** is the recommended first choice. It fixes the bug transparently with zero public API changes and minimal code impact. The performance cost of the parent walk + `EnsureProperties` call is negligible (in-memory walk + one cached server call per `List` instance).

**Approach B** or **C** could be layered on top of A later if there's a concrete need for callers to explicitly provide the list type (e.g., performance-sensitive bulk operations). However, given the negligible cost of auto-detection, this is unlikely to be necessary.

Would love to hear opinions on whether **Approach A alone** is sufficient, or whether the additional overloads from B/C are worth the API surface increase.
