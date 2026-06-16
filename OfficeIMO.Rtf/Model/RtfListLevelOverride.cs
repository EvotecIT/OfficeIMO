namespace OfficeIMO.Rtf;

/// <summary>
/// Per-level override metadata inside an RTF list override table entry.
/// </summary>
public sealed class RtfListLevelOverride {
    /// <summary>Whether this override replaces formatting for the referenced list level.</summary>
    public bool? OverrideFormat { get; set; }

    /// <summary>Whether this override replaces the starting number for the referenced list level.</summary>
    public bool? OverrideStartAt { get; set; }

    /// <summary>Overridden starting number from <c>\levelstartat</c>, when present.</summary>
    public int? StartAt { get; set; }

    /// <summary>Returns whether this override contains any semantic value.</summary>
    public bool HasAnyValue => OverrideFormat.HasValue || OverrideStartAt.HasValue || StartAt.HasValue;
}
