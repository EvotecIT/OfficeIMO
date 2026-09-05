using System;
using OfficeIMO.AsciiDoc.Markdown;
using OfficeIMO.Markdown.Pdf;

namespace OfficeIMO.AsciiDoc.Pdf;

/// <summary>Controls the loss-aware AsciiDoc-to-PDF route.</summary>
public sealed class AsciiDocToPdfOptions {
    private AsciiDocToMarkdownOptions _projectionOptions = new AsciiDocToMarkdownOptions();
    private MarkdownToPdfOptions _markdownOptions = new MarkdownToPdfOptions();

    /// <summary>AsciiDoc-to-Markdown semantic projection settings.</summary>
    public AsciiDocToMarkdownOptions ProjectionOptions {
        get => _projectionOptions;
        set => _projectionOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Shared Markdown/PDF layout, resource, font, and compliance settings.</summary>
    public MarkdownToPdfOptions MarkdownOptions {
        get => _markdownOptions;
        set => _markdownOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    internal AsciiDocToPdfOptions CloneForConversion() => new AsciiDocToPdfOptions {
        ProjectionOptions = new AsciiDocToMarkdownOptions {
            IncludeDocumentAttributesAsFrontMatter = ProjectionOptions.IncludeDocumentAttributesAsFrontMatter,
            PreserveUnsupportedAsSource = ProjectionOptions.PreserveUnsupportedAsSource,
            PreserveCommentsAsSource = ProjectionOptions.PreserveCommentsAsSource,
            ExpandDocumentAttributes = ProjectionOptions.ExpandDocumentAttributes,
            UndefinedAttributeBehavior = ProjectionOptions.UndefinedAttributeBehavior
        },
        MarkdownOptions = MarkdownOptions.Clone()
    };
}
