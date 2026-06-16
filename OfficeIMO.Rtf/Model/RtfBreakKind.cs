namespace OfficeIMO.Rtf;

/// <summary>
/// Type of explicit inline break in an RTF paragraph.
/// </summary>
public enum RtfBreakKind {
    /// <summary>Line break, emitted as <c>\line</c>.</summary>
    Line,

    /// <summary>Soft line break, emitted as <c>\softline</c>.</summary>
    SoftLine,

    /// <summary>Page break, emitted as <c>\page</c>.</summary>
    Page,

    /// <summary>Soft page break, emitted as <c>\softpage</c>.</summary>
    SoftPage,

    /// <summary>Column break, emitted as <c>\column</c>.</summary>
    Column
}
