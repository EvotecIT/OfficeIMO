using OfficeIMO.Markdown.Pdf;
using OfficeIMO.OneNote.Markdown;

namespace OfficeIMO.OneNote.Pdf;

/// <summary>Supported OneNote-to-PDF projection.</summary>
public enum OneNotePdfProjectionMode {
    /// <summary>Projects OneNote into reading-order headings, paragraphs, lists, tables, links, and explicit asset placeholders.</summary>
    SemanticDocument = 0
}

/// <summary>Controls explicit semantic OneNote-to-PDF projection.</summary>
public sealed class OneNotePdfExportOptions {
    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    private OneNoteMarkdownExportOptions _projectionOptions = new OneNoteMarkdownExportOptions();
    private MarkdownPdfExportOptions _markdownOptions = new MarkdownPdfExportOptions();

    /// <summary>Projection mode. OfficeIMO currently supports only the explicit semantic-document contract.</summary>
    public OneNotePdfProjectionMode ProjectionMode { get; set; } = OneNotePdfProjectionMode.SemanticDocument;

    /// <summary>OneNote hierarchy, related-page, and binary-asset projection settings.</summary>
    public OneNoteMarkdownExportOptions ProjectionOptions {
        get => _projectionOptions;
        set => _projectionOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Shared Markdown/PDF layout, resource, font, and compliance settings.</summary>
    public MarkdownPdfExportOptions MarkdownOptions {
        get => _markdownOptions;
        set => _markdownOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    internal OneNotePdfExportOptions CloneForConversion() {
        if (ProjectionMode != OneNotePdfProjectionMode.SemanticDocument) {
            throw new ArgumentOutOfRangeException(nameof(ProjectionMode), ProjectionMode, "Unsupported OneNote PDF projection mode.");
        }
        return new OneNotePdfExportOptions {
            ProjectionMode = ProjectionMode,
            ProjectionOptions = ProjectionOptions.Clone(),
            MarkdownOptions = MarkdownOptions.Clone()
        };
    }

    /// <summary>Creates an independent copy of these settings.</summary>
    public OneNotePdfExportOptions Clone() {
        OneNotePdfExportOptions copy = CloneForConversion();
        copy.FailOnLoss = FailOnLoss;
        return copy;
    }
}
