namespace OfficeIMO.Rtf;

/// <summary>
/// Level kind for Word 6/95 legacy RTF paragraph numbering.
/// </summary>
public enum RtfLegacyNumberingLevelKind {
    /// <summary>No explicit legacy numbering level kind.</summary>
    None,

    /// <summary>Explicit numbered paragraph level represented by <c>\pnlvl</c>.</summary>
    Level,

    /// <summary>Bulleted paragraph represented by <c>\pnlvlblt</c>.</summary>
    Bullet,

    /// <summary>Simple numbered paragraph represented by <c>\pnlvlbody</c>.</summary>
    Body,

    /// <summary>Continue numbering without displaying a number, represented by <c>\pnlvlcont</c>.</summary>
    Continue
}
