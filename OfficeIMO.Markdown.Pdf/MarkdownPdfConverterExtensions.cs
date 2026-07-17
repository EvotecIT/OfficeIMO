using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    /// <summary>
    /// Converts a Markdown document model to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDocument ToPdfDocument(this MarkdownDoc document, MarkdownPdfSaveOptions? options = null) {
        return document.ToPdfDocumentResult(options).Value;
    }

    internal static PdfCore.PdfDocument ConvertToPdfDocument(MarkdownDoc document, MarkdownPdfSaveOptions options) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        PdfCore.PdfOptions pdfOptions = options.PdfOptions?.Clone() ?? new PdfCore.PdfOptions();
        pdfOptions.ReportDiagnosticsTo(options.Report, "OfficeIMO.Markdown.Pdf");

        ApplyMarkdownTextFallbackOptions(pdfOptions, options, document);

        if (options.CreateOutlineFromHeadings) {
            pdfOptions.CreateOutlineFromHeadings = true;
        }

        MarkdownPdfStyle visualTheme = ResolveVisualTheme(document, options);
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(pdfOptions);
        PdfCore.PdfTheme? documentTheme = visualTheme.DocumentThemeSnapshot;
        if (documentTheme != null) {
            pdf.Theme(documentTheme);
        }
        ApplyMarkdownDefaultFont(pdf, options);
        visualTheme.ApplyPageDecorations(pdf, pdfOptions);

        IReadOnlyList<IMarkdownBlock> topLevelBlocks = GetPdfTopLevelBlocks(document);
        ApplyMetadata(pdf, document, options);
        string? promotedFrontMatterTitle = GetPromotedFrontMatterTitle(document, options);
        RenderBlocks(pdf, topLevelBlocks, document, options, visualTheme, promotedFrontMatterTitle);
        if (topLevelBlocks.Count == 0) {
            pdf.Paragraph(paragraph => paragraph.Text(string.Empty));
        }

        return pdf;
    }

    /// <summary>
    /// Converts a Markdown document model to a PDF document and returns conversion diagnostics with it.
    /// </summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this MarkdownDoc document, MarkdownPdfSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        MarkdownPdfSaveOptions operation = (options ?? new MarkdownPdfSaveOptions()).CloneForConversion();
        PdfCore.PdfDocument pdf = ConvertToPdfDocument(document, operation);
        return new PdfCore.PdfDocumentConversionResult(pdf, operation.Report);
    }

    private static void ApplyMarkdownDefaultFont(PdfCore.PdfDocument pdf, MarkdownPdfSaveOptions options) {
        if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(options.FontFamily, out PdfCore.PdfStandardFont font)) {
            pdf.DefaultTextStyle(style => style.Font(PdfCore.PdfStandardFontMapper.GetFontFamily(font)));
        }
    }

    private static IReadOnlyList<IMarkdownBlock> GetPdfTopLevelBlocks(MarkdownDoc document) {
        var (blocks, _) = document.GetBlocksAndHeadingSlugs();
        if (document.DocumentHeader == null) {
            return blocks;
        }

        var withFrontMatter = new List<IMarkdownBlock>(blocks.Count + 1) {
            document.DocumentHeader
        };
        withFrontMatter.AddRange(blocks);
        return withFrontMatter;
    }

    /// <summary>
    /// Converts a Markdown document model to PDF bytes.
    /// </summary>
    /// <example><code>byte[] pdf = document.ToPdf();</code></example>
    public static byte[] ToPdf(this MarkdownDoc document, MarkdownPdfSaveOptions? options = null) {
        return document.ToPdfDocument(options).ToBytes();
    }

    /// <summary>
    /// Saves a Markdown document model as a PDF file.
    /// </summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(this MarkdownDoc document, string path, MarkdownPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(path);

    /// <summary>
    /// Attempts to save a Markdown document model as a PDF file and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this MarkdownDoc document, string path, MarkdownPdfSaveOptions? options = null) {
        try {
            return document.ToPdfDocumentResult(options).TrySave(path);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>
    /// Writes a Markdown document model as PDF to a stream.
    /// </summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(this MarkdownDoc document, Stream stream, MarkdownPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(stream);

    /// <summary>
    /// Attempts to write a Markdown document model as PDF to a stream and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this MarkdownDoc document, Stream stream, MarkdownPdfSaveOptions? options = null) {
        try {
            return document.ToPdfDocumentResult(options).TrySave(stream);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    /// <summary>Converts synchronously, then asynchronously saves a Markdown PDF at the specified path.</summary>
    public static Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(
        this MarkdownDoc document,
        string path,
        MarkdownPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return document.ToPdfDocumentResult(options).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously saves a Markdown PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(
        this MarkdownDoc document,
        Stream stream,
        MarkdownPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return document.ToPdfDocumentResult(options).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save a Markdown document model as PDF at the specified path.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
        this MarkdownDoc document,
        string path,
        MarkdownPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await document.ToPdfDocumentResult(options)
                .TrySaveAsync(path, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>Attempts to asynchronously save a Markdown document model as PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
        this MarkdownDoc document,
        Stream stream,
        MarkdownPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await document.ToPdfDocumentResult(options)
                .TrySaveAsync(stream, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    private static MarkdownPdfStyle ResolveVisualTheme(MarkdownDoc document, MarkdownPdfSaveOptions options) {
        MarkdownPdfStyle? explicitTheme = options.Style;
        if (explicitTheme != null) {
            return explicitTheme;
        }

        MarkdownVisualTheme? sharedTheme = options.ThemeSnapshot;
        if (sharedTheme != null) {
            return MarkdownPdfStyle.FromMarkdownTheme(sharedTheme);
        }

        if (options.UseFrontMatterTheme && document.DocumentHeader != null) {
            string? frontMatterPdfTheme = GetFrontMatterMetadata(document.DocumentHeader, "pdfTheme")
                ?? GetFrontMatterMetadata(document.DocumentHeader, "pdf-theme");
            if (frontMatterPdfTheme != null) {
                if (TryResolveTheme(frontMatterPdfTheme, out MarkdownPdfStyle? theme)) {
                    return theme!;
                }

                AddWarning(options, "UnsupportedVisualTheme", frontMatterPdfTheme, "The requested Markdown PDF visual theme is not recognized; the configured fallback visual profile is used.");
            }

            string? frontMatterTheme = GetFrontMatterMetadata(document.DocumentHeader, "theme")
                ?? GetFrontMatterMetadata(document.DocumentHeader, "visualTheme")
                ?? GetFrontMatterMetadata(document.DocumentHeader, "visual-theme");
            if (frontMatterTheme != null) {
                if (TryResolveTheme(frontMatterTheme, out MarkdownPdfStyle? theme)) {
                    return theme!;
                }

                AddWarning(options, "UnsupportedVisualTheme", frontMatterTheme, "The requested Markdown visual theme is not recognized; the configured fallback visual profile is used.");
            }
        }

        MarkdownVisualTheme? defaultTheme = MarkdownVisualTheme.ResolveOrDefault(null, options.ApplyDefaultTheme);
        return defaultTheme != null
            ? MarkdownPdfStyle.FromMarkdownTheme(defaultTheme)
            : MarkdownPdfStyle.Plain();
    }

    private static bool TryResolveTheme(string themeName, out MarkdownPdfStyle? theme) {
        theme = null;
        if (MarkdownVisualTheme.TryCreate(themeName, out MarkdownVisualTheme? markdownTheme)) {
            theme = MarkdownPdfStyle.FromMarkdownTheme(markdownTheme!);
            return true;
        }
        return false;
    }

    private static void RenderBlocks(PdfCore.PdfDocument pdf, IEnumerable<IMarkdownBlock> blocks, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme, string? skipFirstHeadingTitle = null) {
        bool skippedPromotedHeading = false;
        var materializedBlocks = blocks as IReadOnlyList<IMarkdownBlock> ?? blocks.ToList();
        for (int i = 0; i < materializedBlocks.Count; i++) {
            IMarkdownBlock block = materializedBlocks[i];
            if (!skippedPromotedHeading && skipFirstHeadingTitle != null && block is HeadingBlock heading && heading.Level == 1 && IsSameNormalizedText(heading.Text, skipFirstHeadingTitle)) {
                skippedPromotedHeading = true;
                continue;
            }

            if (block is HeadingBlock tocTitleHeading &&
                i + 1 < materializedBlocks.Count &&
                materializedBlocks[i + 1] is TocBlock toc &&
                ShouldRenderTocAsPanel(toc) &&
                toc.IncludeTitle &&
                IsSameNormalizedText(tocTitleHeading.Text, toc.Title)) {
                continue;
            }

            RenderBlock(pdf, block, document, options, visualTheme);
        }
    }

    private static void ApplyMetadata(PdfCore.PdfDocument pdf, MarkdownDoc document, MarkdownPdfSaveOptions options) {
        string? title = NormalizeMetadata(options.Title);
        string? author = NormalizeMetadata(options.Author);
        string? subject = NormalizeMetadata(options.Subject);
        string? keywords = NormalizeMetadata(options.Keywords);

        if (options.UseFrontMatterMetadata) {
            FrontMatterBlock? frontMatter = document.DocumentHeader;
            if (frontMatter != null) {
                title ??= GetFrontMatterMetadata(frontMatter, "title");
                author ??= GetFrontMatterMetadata(frontMatter, "author");
                subject ??= GetFrontMatterMetadata(frontMatter, "subject") ?? GetFrontMatterMetadata(frontMatter, "description") ?? GetFrontMatterMetadata(frontMatter, "summary");
                keywords ??= GetFrontMatterMetadata(frontMatter, "keywords") ?? GetFrontMatterMetadata(frontMatter, "tags");
            }
        }

        if (title == null && options.UseFirstHeadingAsTitle) {
            title = NormalizeMetadata(document.Blocks.OfType<HeadingBlock>().FirstOrDefault()?.Text);
        }

        if (title != null || author != null || subject != null || keywords != null) {
            pdf.Meta(title, author, subject, keywords);
        }
    }

    private static string? GetFrontMatterMetadata(FrontMatterBlock frontMatter, string key) {
        FrontMatterBlock.Entry? entry = frontMatter.FindEntry(key);
        return entry == null ? null : NormalizeMetadata(ConvertMetadataValue(entry.Value));
    }

    private static string? GetPromotedFrontMatterTitle(MarkdownDoc document, MarkdownPdfSaveOptions options) {
        if (options.FrontMatterRenderMode != MarkdownPdfFrontMatterRenderMode.DocumentHeader || document.DocumentHeader == null) {
            return null;
        }

        return GetFrontMatterMetadata(document.DocumentHeader, "title");
    }

    private static string? ConvertMetadataValue(object? value) {
        switch (value) {
            case null:
                return null;
            case string text:
                return text;
            case IEnumerable<string> values:
                return string.Join(", ", values.Where(item => !string.IsNullOrWhiteSpace(item)).Select(item => item.Trim()));
            case System.Collections.IEnumerable values:
                var items = new List<string>();
                foreach (object? item in values) {
                    string? normalized = NormalizeMetadata(Convert.ToString(item, CultureInfo.InvariantCulture));
                    if (normalized != null) {
                        items.Add(normalized);
                    }
                }

                return items.Count == 0 ? null : string.Join(", ", items);
            default:
                return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
    }
}
