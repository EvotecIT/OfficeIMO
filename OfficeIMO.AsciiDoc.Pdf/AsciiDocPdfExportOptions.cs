using System;
using OfficeIMO.AsciiDoc.Markdown;
using OfficeIMO.Markdown.Pdf;

namespace OfficeIMO.AsciiDoc.Pdf;

/// <summary>Controls the loss-aware AsciiDoc-to-PDF route.</summary>
public sealed class AsciiDocPdfExportOptions {
    private AsciiDocMarkdownExportOptions _projectionOptions = new AsciiDocMarkdownExportOptions();
    private MarkdownPdfExportOptions _markdownOptions = new MarkdownPdfExportOptions();

    /// <summary>AsciiDoc-to-Markdown semantic projection settings.</summary>
    public AsciiDocMarkdownExportOptions ProjectionOptions {
        get => _projectionOptions;
        set => _projectionOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Shared Markdown/PDF layout, resource, font, and compliance settings.</summary>
    public MarkdownPdfExportOptions MarkdownOptions {
        get => _markdownOptions;
        set => _markdownOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    internal AsciiDocPdfExportOptions CloneForConversion() => new AsciiDocPdfExportOptions {
        ProjectionOptions = new AsciiDocMarkdownExportOptions {
            IncludeDocumentAttributesAsFrontMatter = ProjectionOptions.IncludeDocumentAttributesAsFrontMatter,
            PreserveUnsupportedAsSource = ProjectionOptions.PreserveUnsupportedAsSource,
            PreserveCommentsAsSource = ProjectionOptions.PreserveCommentsAsSource,
            ExpandDocumentAttributes = ProjectionOptions.ExpandDocumentAttributes,
            UndefinedAttributeBehavior = ProjectionOptions.UndefinedAttributeBehavior
        },
        MarkdownOptions = MarkdownOptions.Clone()
    };
}
