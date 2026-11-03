using OfficeIMO.Markdown;
using OfficeIMO.OneNote.Markdown;

namespace OfficeIMO.OneNote.Html;

/// <summary>Controls semantic OneNote-to-HTML conversion through the shared Markdown model.</summary>
public sealed class OneNoteHtmlExportOptions {
    /// <summary>OneNote-to-Markdown projection settings. Null uses the defaults.</summary>
    public OneNoteMarkdownExportOptions? ProjectionOptions { get; set; }

    /// <summary>Markdown-to-HTML rendering settings. Null uses the defaults.</summary>
    public MarkdownHtmlExportOptions? HtmlOptions { get; set; }

    /// <summary>Produces an embeddable fragment instead of a standalone HTML5 document.</summary>
    public bool Fragment { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates an independent copy of these settings.</summary>
    public OneNoteHtmlExportOptions Clone() => new OneNoteHtmlExportOptions {
        ProjectionOptions = ProjectionOptions?.Clone(),
        HtmlOptions = HtmlOptions?.CloneForRender(),
        Fragment = Fragment,
        FailOnLoss = FailOnLoss
    };
}
