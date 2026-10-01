namespace OfficeIMO.Rtf;

/// <summary>
/// Explicit inline break in an RTF paragraph.
/// </summary>
public sealed partial class RtfBreak : IRtfInline {
    /// <summary>Creates a break of the specified kind.</summary>
    public RtfBreak(RtfBreakKind kind) {
        Kind = kind;
    }

    /// <summary>Break kind.</summary>
    public RtfBreakKind Kind { get; set; }
}
