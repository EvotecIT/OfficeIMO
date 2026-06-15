namespace OfficeIMO.Rtf;

/// <summary>
/// RTF list override entry that maps <c>\ls</c> paragraph ids to list definitions.
/// </summary>
public sealed class RtfListOverride {
    private readonly List<RtfListLevelOverride> _levelOverrides = new List<RtfListLevelOverride>();

    /// <summary>Creates a list override.</summary>
    public RtfListOverride(int id, int listId) {
        Id = id;
        ListId = listId;
    }

    /// <summary>Paragraph list override id used by <c>\ls</c>.</summary>
    public int Id { get; }

    /// <summary>Referenced list definition id.</summary>
    public int ListId { get; }

    /// <summary>Number of level overrides, when present.</summary>
    public int? OverrideCount { get; set; }

    /// <summary>Per-level override metadata in declaration order.</summary>
    public IReadOnlyList<RtfListLevelOverride> LevelOverrides => _levelOverrides.AsReadOnly();

    /// <summary>Adds a per-level override entry.</summary>
    public RtfListLevelOverride AddLevelOverride() {
        var levelOverride = new RtfListLevelOverride();
        _levelOverrides.Add(levelOverride);
        return levelOverride;
    }

    internal void AddParsedLevelOverride(RtfListLevelOverride levelOverride) {
        _levelOverrides.Add(levelOverride ?? throw new ArgumentNullException(nameof(levelOverride)));
    }
}
