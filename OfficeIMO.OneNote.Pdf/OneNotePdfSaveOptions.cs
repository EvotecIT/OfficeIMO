using OfficeIMO.Markdown.Pdf;
using OfficeIMO.OneNote.Markdown;

namespace OfficeIMO.OneNote.Pdf;

/// <summary>Supported OneNote-to-PDF projection.</summary>
public enum OneNotePdfProjectionMode {
    /// <summary>Projects OneNote into reading-order headings, paragraphs, lists, tables, links, and explicit asset placeholders.</summary>
    SemanticDocument = 0
}

/// <summary>Controls explicit semantic OneNote-to-PDF projection.</summary>
public sealed class OneNotePdfSaveOptions {
    private OneNoteMarkdownOptions _projectionOptions = new OneNoteMarkdownOptions();
    private MarkdownPdfSaveOptions _pdfOptions = new MarkdownPdfSaveOptions();

    /// <summary>Projection mode. OfficeIMO currently supports only the explicit semantic-document contract.</summary>
    public OneNotePdfProjectionMode ProjectionMode { get; set; } = OneNotePdfProjectionMode.SemanticDocument;

    /// <summary>OneNote hierarchy, related-page, and binary-asset projection settings.</summary>
    public OneNoteMarkdownOptions ProjectionOptions {
        get => _projectionOptions;
        set => _projectionOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Shared Markdown/PDF layout, resource, font, and compliance settings.</summary>
    public MarkdownPdfSaveOptions PdfOptions {
        get => _pdfOptions;
        set => _pdfOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    internal OneNotePdfSaveOptions CloneForConversion() {
        if (ProjectionMode != OneNotePdfProjectionMode.SemanticDocument) {
            throw new ArgumentOutOfRangeException(nameof(ProjectionMode), ProjectionMode, "Unsupported OneNote PDF projection mode.");
        }
        return new OneNotePdfSaveOptions {
            ProjectionMode = ProjectionMode,
            ProjectionOptions = ProjectionOptions.Clone(),
            PdfOptions = PdfOptions.Clone()
        };
    }
}
