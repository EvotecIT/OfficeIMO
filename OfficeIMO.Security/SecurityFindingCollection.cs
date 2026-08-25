namespace OfficeIMO.Security;

/// <summary>Collects findings without allocating a list for a successful validation.</summary>
internal struct SecurityFindingCollection {
    private List<SecurityFinding>? _items;

    internal readonly IReadOnlyList<SecurityFinding> Items =>
        _items ?? (IReadOnlyList<SecurityFinding>)Array.Empty<SecurityFinding>();

    internal void Add(SecurityFinding finding) => (_items ??= new List<SecurityFinding>()).Add(finding);

    internal readonly SecurityFinding[] ToArray() =>
        _items == null ? Array.Empty<SecurityFinding>() : _items.ToArray();
}
