using System;
using OfficeIMO.Latex.Markdown;
using OfficeIMO.Markdown.Pdf;

namespace OfficeIMO.Latex.Pdf;

/// <summary>Controls the loss-aware LaTeX-to-PDF route.</summary>
public sealed class LatexPdfSaveOptions {
    private LatexToMarkdownOptions _projectionOptions = new LatexToMarkdownOptions();
    private MarkdownPdfSaveOptions _pdfOptions = new MarkdownPdfSaveOptions();

    /// <summary>LaTeX-to-Markdown semantic projection settings.</summary>
    public LatexToMarkdownOptions ProjectionOptions {
        get => _projectionOptions;
        set => _projectionOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Shared Markdown/PDF layout, resource, font, and compliance settings.</summary>
    public MarkdownPdfSaveOptions PdfOptions {
        get => _pdfOptions;
        set => _pdfOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    internal LatexPdfSaveOptions CloneForConversion() => new LatexPdfSaveOptions {
        ProjectionOptions = new LatexToMarkdownOptions {
            PreserveUnsupportedAsSource = ProjectionOptions.PreserveUnsupportedAsSource,
            IncludePreambleAsFrontMatter = ProjectionOptions.IncludePreambleAsFrontMatter
        },
        PdfOptions = PdfOptions.Clone()
    };
}
