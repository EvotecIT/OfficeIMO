namespace OfficeIMO.Pdf;

/// <summary>
/// Catalog-level tagged PDF structure metadata read from /MarkInfo and /StructTreeRoot.
/// </summary>
public sealed class PdfTaggedContentInfo {
    internal PdfTaggedContentInfo(
        bool? marked,
        bool? suspects,
        bool? userProperties,
        int? structTreeRootObjectNumber,
        int? parentTreeObjectNumber,
        int? parentTreeNextKey,
        IReadOnlyDictionary<string, string> roleMap,
        IReadOnlyList<int> rootElementObjectNumbers,
        IReadOnlyList<int> parentTreeStructParentIndexes,
        IReadOnlyList<PdfStructureElementInfo> structureElements) {
        Marked = marked;
        Suspects = suspects;
        UserProperties = userProperties;
        StructTreeRootObjectNumber = structTreeRootObjectNumber;
        ParentTreeObjectNumber = parentTreeObjectNumber;
        ParentTreeNextKey = parentTreeNextKey;
        RoleMap = roleMap;
        RootElementObjectNumbers = rootElementObjectNumbers;
        ParentTreeStructParentIndexes = parentTreeStructParentIndexes;
        StructureElements = structureElements;
    }

    /// <summary>Catalog /MarkInfo /Marked value, when present.</summary>
    public bool? Marked { get; }

    /// <summary>Catalog /MarkInfo /Suspects value, when present.</summary>
    public bool? Suspects { get; }

    /// <summary>Catalog /MarkInfo /UserProperties value, when present.</summary>
    public bool? UserProperties { get; }

    /// <summary>StructTreeRoot object number, when the catalog references one indirectly.</summary>
    public int? StructTreeRootObjectNumber { get; }

    /// <summary>ParentTree object number, when the structure tree references one indirectly.</summary>
    public int? ParentTreeObjectNumber { get; }

    /// <summary>StructTreeRoot /ParentTreeNextKey value, when present.</summary>
    public int? ParentTreeNextKey { get; }

    /// <summary>Role map entries from /RoleMap, keyed by custom role name.</summary>
    public IReadOnlyDictionary<string, string> RoleMap { get; }

    /// <summary>True when at least one role-map entry was readable.</summary>
    public bool HasRoleMap => RoleMap.Count > 0;

    /// <summary>Top-level structure element object numbers referenced by StructTreeRoot /K.</summary>
    public IReadOnlyList<int> RootElementObjectNumbers { get; }

    /// <summary>StructParent indexes discovered in the ParentTree /Nums array.</summary>
    public IReadOnlyList<int> ParentTreeStructParentIndexes { get; }

    /// <summary>Number of ParentTree /Nums entries discovered.</summary>
    public int ParentTreeEntryCount => ParentTreeStructParentIndexes.Count;

    /// <summary>Structure elements discovered by scanning reachable parsed objects.</summary>
    public IReadOnlyList<PdfStructureElementInfo> StructureElements { get; }

    /// <summary>Number of readable structure element objects.</summary>
    public int StructureElementCount => StructureElements.Count;

    /// <summary>Distinct readable structure types in first-seen object order.</summary>
    public IReadOnlyList<string> StructureTypes => StructureElements
        .Select(element => element.StructureType)
        .Where(type => !string.IsNullOrEmpty(type))
        .Cast<string>()
        .Distinct(StringComparer.Ordinal)
        .ToArray();
}
