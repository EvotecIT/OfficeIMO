using System;
using OfficeIMO.AsciiDoc.Markdown;
using OfficeIMO.Markdown.Pdf;

namespace OfficeIMO.AsciiDoc.Pdf;

/// <summary>Controls the loss-aware AsciiDoc-to-PDF route.</summary>
public sealed class AsciiDocPdfSaveOptions {
    private AsciiDocToMarkdownOptions _projectionOptions = new AsciiDocToMarkdownOptions();
    private MarkdownPdfSaveOptions _pdfOptions = new MarkdownPdfSaveOptions();

    /// <summary>AsciiDoc-to-Markdown semantic projection settings.</summary>
    public AsciiDocToMarkdownOptions ProjectionOptions {
        get => _projectionOptions;
        set => _projectionOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Shared Markdown/PDF layout, resource, font, and compliance settings.</summary>
    public MarkdownPdfSaveOptions PdfOptions {
        get => _pdfOptions;
        set => _pdfOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    internal AsciiDocPdfSaveOptions CloneForConversion() => new AsciiDocPdfSaveOptions {
        ProjectionOptions = new AsciiDocToMarkdownOptions {
            IncludeDocumentAttributesAsFrontMatter = ProjectionOptions.IncludeDocumentAttributesAsFrontMatter,
            PreserveUnsupportedAsSource = ProjectionOptions.PreserveUnsupportedAsSource,
            PreserveCommentsAsSource = ProjectionOptions.PreserveCommentsAsSource,
            ExpandDocumentAttributes = ProjectionOptions.ExpandDocumentAttributes,
            UndefinedAttributeBehavior = ProjectionOptions.UndefinedAttributeBehavior
        },
        PdfOptions = PdfOptions.Clone()
    };
}
