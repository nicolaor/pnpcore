# Folder Creation for Non-Document-Library Lists — Proposed Approaches

## Context

The current fix in the `Folder()` constructor's `AddApiCallHandler` uses `FindParentList()` to walk the in-memory parent hierarchy, find the owning `List`, then call `EnsurePropertiesAsync(p => p.BaseType)` to decide between `Folders/AddUsingPath` (document libraries) and `AddSubFolderUsingPath` (generic lists). This works transparently but involves a parent-pointer walk and a potential extra server call.

Below are three alternative approaches that let the **caller** provide the list type directly, all of which preserve backward compatibility (existing callers continue to work without changes).

---

## Approach A — New overloads with an optional `ListBaseType?` parameter

Add new overloads alongside the existing ones. Existing signatures stay untouched, so nothing breaks.

### Files changed

| File | Change |
|------|--------|
| `IFolderCollection.cs` | Add 4 new overloads with `ListBaseType? listBaseType = null` |
| `FolderCollection.cs` | Implement the new overloads, pass `listBaseType` through `keyValuePairs` dictionary |
| `IFolder.cs` | Add 4 new overloads for `AddFolderAsync`/`AddFolderBatch*` with `ListBaseType?` |
| `Folder.cs` (extension methods) | Implement the new overloads, forward to `FolderCollection` |
| `Folder.cs` (constructor) | Read `listBaseType` from `keyValuePairs` when present; fall back to `FindParentList` when `null` |

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

    // 2. Fall back to auto-detection via parent walk (current behavior)
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
❌ Adds ~8 overloads across `IFolderCollection`/`IFolder`/`FolderCollection`/`Folder`.

---

## Approach B — `FolderAddOptions` class (more extensible)

Introduce an options object (similar to `MoveCopyOptions` used by `CopyToAsync`/`MoveToAsync`) and add new overloads that accept it.

### Files changed

| File | Change |
|------|--------|
| New `FolderAddOptions.cs` | New class with `ListBaseType?` property |
| `IFolderCollection.cs` | Add 4 new overloads with `FolderAddOptions options = null` |
| `FolderCollection.cs` | Implement new overloads, pack options into `keyValuePairs` |
| `IFolder.cs` | Add 4 new `AddFolderAsync`/`AddFolderBatch*` overloads |
| `Folder.cs` (extension methods + constructor) | Same as Approach A for the handler logic |

### Code sketch — `FolderAddOptions.cs`

```csharp
public class FolderAddOptions
{
    /// <summary>
    /// Hint for the base type of the parent list. When provided, skips automatic detection
    /// via the parent model hierarchy. When null, the SDK auto-detects as before.
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

## Approach C — Keep the current auto-detection (`FindParentList`), do nothing extra

The current implementation already works transparently for all callers. The costs are minimal:

- `FindParentList` — in-memory parent-pointer walk, zero server calls
- `EnsurePropertiesAsync(BaseType)` — one server call per `List` instance, cached after first invocation; creating N folders under the same list incurs 0 extra round-trips after the first

No public API changes. No files changed.

### Caller usage — same as today

```csharp
await list.RootFolder.Folders.AddAsync("MyFolder");  // just works
await list.RootFolder.EnsureFolderAsync("sub1/sub2"); // just works
```

### Pros & Cons

✅ No new API surface; zero burden on callers; already implemented.
❌ Caller can't skip the `FindParentList` + `EnsureProperties` code path (though the cost is negligible).

---

## Recommendation

**Approach A** (or B) layered on top of C as a "fast path" seems ideal — the auto-detection stays as the default, but callers who already know the list type can skip it. However, given the negligible cost of auto-detection, **Approach C alone** may be perfectly sufficient.

Would love to hear opinions on whether the additional overloads are worth the API surface increase, or whether auto-detection is "good enough" for all practical scenarios.
