using System;
using OfficeIMO.Latex.Markdown;
using OfficeIMO.Markdown.Pdf;

namespace OfficeIMO.Latex.Pdf;

/// <summary>Controls the loss-aware LaTeX-to-PDF route.</summary>
public sealed class LatexPdfExportOptions {
    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    private LatexMarkdownExportOptions _projectionOptions = new LatexMarkdownExportOptions();
    private MarkdownPdfExportOptions _markdownOptions = new MarkdownPdfExportOptions();

    /// <summary>LaTeX-to-Markdown semantic projection settings.</summary>
    public LatexMarkdownExportOptions ProjectionOptions {
        get => _projectionOptions;
        set => _projectionOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Shared Markdown/PDF layout, resource, font, and compliance settings.</summary>
    public MarkdownPdfExportOptions MarkdownOptions {
        get => _markdownOptions;
        set => _markdownOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    internal LatexPdfExportOptions CloneForConversion() => new LatexPdfExportOptions {
        ProjectionOptions = new LatexMarkdownExportOptions {
            PreserveUnsupportedAsSource = ProjectionOptions.PreserveUnsupportedAsSource,
            IncludePreambleAsFrontMatter = ProjectionOptions.IncludePreambleAsFrontMatter
        },
        MarkdownOptions = MarkdownOptions.Clone()
    };

    /// <summary>Creates an independent copy of these settings.</summary>
    public LatexPdfExportOptions Clone() {
        LatexPdfExportOptions copy = CloneForConversion();
        copy.FailOnLoss = FailOnLoss;
        return copy;
    }
}
