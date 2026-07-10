namespace OfficeIMO.AsciiDoc;

/// <summary>AsciiDoc writer behavior.</summary>
public enum AsciiDocWriterMode {
    /// <summary>Reuse exact original source for unchanged subtrees.</summary>
    Preserve = 0,
    /// <summary>Emit stable OfficeIMO formatting for recognized nodes.</summary>
    Canonical = 1
}
/// <summary>Options controlling AsciiDoc writing.</summary>
public sealed class AsciiDocWriterOptions {
    /// <summary>Writer mode. Defaults to source preservation.</summary>
    public AsciiDocWriterMode Mode { get; set; } = AsciiDocWriterMode.Preserve;

    /// <summary>Canonical output line ending. Null uses the source document's preferred line ending.</summary>
    public string? LineEnding { get; set; }
}
