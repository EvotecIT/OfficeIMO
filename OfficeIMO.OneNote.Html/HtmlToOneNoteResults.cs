using OfficeIMO.Html;

namespace OfficeIMO.OneNote.Html;

/// <summary>Typed OneNote section plus structured HTML-import evidence.</summary>
public sealed class HtmlToOneNoteSectionResult : HtmlConversionResult<OneNoteSection> {
    internal HtmlToOneNoteSectionResult(OneNoteSection value) : base(value) { }
    internal void AddImportDiagnostic(HtmlDiagnostic diagnostic) => AddDiagnostic(diagnostic);
    /// <summary>Number of imported pages.</summary>
    public int Pages { get; internal set; }
    /// <summary>Number of imported native content elements.</summary>
    public int Elements { get; internal set; }
    /// <summary>Number of imported tables.</summary>
    public int Tables { get; internal set; }
    /// <summary>Number of imported data URI images.</summary>
    public int Images { get; internal set; }
}

/// <summary>Typed OneNote notebook plus structured HTML-import evidence.</summary>
public sealed class HtmlToOneNoteNotebookResult : HtmlConversionResult<OneNoteNotebook> {
    internal HtmlToOneNoteNotebookResult(OneNoteNotebook value) : base(value) { }
    internal void AddImportDiagnostic(HtmlDiagnostic diagnostic) => AddDiagnostic(diagnostic);
    /// <summary>Number of imported sections.</summary>
    public int Sections { get; internal set; }
    /// <summary>Number of imported pages.</summary>
    public int Pages { get; internal set; }
    /// <summary>Number of imported native content elements.</summary>
    public int Elements { get; internal set; }
    /// <summary>Number of imported tables.</summary>
    public int Tables { get; internal set; }
    /// <summary>Number of imported data URI images.</summary>
    public int Images { get; internal set; }
}
