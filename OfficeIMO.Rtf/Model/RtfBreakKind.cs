namespace OfficeIMO.Rtf;

/// <summary>
/// Type of explicit inline break in an RTF paragraph.
/// </summary>
public enum RtfBreakKind {
    /// <summary>Line break, emitted as <c>\line</c>.</summary>
    Line,

    /// <summary>Page break, emitted as <c>\page</c>.</summary>
    Page,

    /// <summary>Column break, emitted as <c>\column</c>.</summary>
    Column
}
