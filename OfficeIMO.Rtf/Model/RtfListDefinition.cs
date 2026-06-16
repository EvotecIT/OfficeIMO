namespace OfficeIMO.Rtf;

/// <summary>
/// RTF list definition from the list table.
/// </summary>
public sealed class RtfListDefinition {
    private readonly List<RtfListLevel> _levels = new List<RtfListLevel>();

    /// <summary>Creates a list definition.</summary>
    public RtfListDefinition(int id) {
        Id = id;
    }

    /// <summary>List definition id referenced by list overrides.</summary>
    public int Id { get; }

    /// <summary>Optional template id.</summary>
    public int? TemplateId { get; set; }

    /// <summary>Optional list name.</summary>
    public string? Name { get; set; }

    /// <summary>List levels in definition order.</summary>
    public IReadOnlyList<RtfListLevel> Levels => _levels.AsReadOnly();

    /// <summary>Adds a list level.</summary>
    public RtfListLevel AddLevel(RtfListKind kind = RtfListKind.Decimal) {
        var level = new RtfListLevel(_levels.Count, kind);
        _levels.Add(level);
        return level;
    }

    internal void AddParsedLevel(RtfListLevel level) {
        _levels.Add(level ?? throw new ArgumentNullException(nameof(level)));
    }
}
