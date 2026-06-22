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

    /// <summary>True when a readable top-level /Document structure element was discovered.</summary>
    public bool HasDocumentStructureElement => StructureElements.Any(static element => string.Equals(element.StructureType, "Document", StringComparison.Ordinal));

    /// <summary>Total marked-content references discovered in readable structure element /K entries.</summary>
    public int MarkedContentReferenceCount => StructureElements.Sum(static element => element.MarkedContentReferenceCount);

    /// <summary>Total object references discovered in readable structure element /K entries.</summary>
    public int ObjectReferenceCount => StructureElements.Sum(static element => element.ObjectReferenceCount);

    /// <summary>True when at least one structure element links to marked page content.</summary>
    public bool HasMarkedContentReferences => MarkedContentReferenceCount > 0;

    /// <summary>True when at least one structure element references an object such as a widget or annotation.</summary>
    public bool HasObjectReferences => ObjectReferenceCount > 0;

    /// <summary>Distinct readable structure types in first-seen object order.</summary>
    public IReadOnlyList<string> StructureTypes => StructureElements
        .Select(element => element.StructureType)
        .Where(type => !string.IsNullOrEmpty(type))
        .Cast<string>()
        .Distinct(StringComparer.Ordinal)
        .ToArray();

    /// <summary>Readable structure element counts grouped by structure type.</summary>
    public IReadOnlyDictionary<string, int> StructureTypeCounts => StructureElements
        .Where(static element => !string.IsNullOrEmpty(element.StructureType))
        .GroupBy(static element => element.StructureType!, StringComparer.Ordinal)
        .ToDictionary(static group => group.Key, static group => group.Count(), StringComparer.Ordinal);

    /// <summary>Number of structure elements with a readable /Lang value.</summary>
    public int LanguageElementCount => StructureElements.Count(static element => !string.IsNullOrEmpty(element.Language));

    /// <summary>Number of structure elements with readable /Alt alternate text.</summary>
    public int AlternateTextElementCount => StructureElements.Count(static element => !string.IsNullOrEmpty(element.AlternateText));

    /// <summary>Number of Figure structure elements without readable alternate text.</summary>
    public int FigureWithoutAlternateTextCount => StructureElements.Count(static element =>
        string.Equals(element.StructureType, "Figure", StringComparison.Ordinal) &&
        string.IsNullOrEmpty(element.AlternateText));

    /// <summary>True when role mapping, language, alternate text, and parent-tree evidence is present for deeper tagged PDF workflows.</summary>
    public bool HasDeepTaggedPdfEvidence =>
        Marked == true &&
        StructureElementCount > 0 &&
        ParentTreeEntryCount > 0 &&
        (HasRoleMap || StructureTypeCounts.Count > 0);

    /// <summary>True when all readable Figure structure elements have alternate text.</summary>
    public bool FiguresHaveAlternateText => FigureWithoutAlternateTextCount == 0;
}
