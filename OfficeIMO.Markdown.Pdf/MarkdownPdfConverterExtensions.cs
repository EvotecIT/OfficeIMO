using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    /// <summary>
    /// Converts Markdown text to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDocument ToPdfDocument(this string markdown, MarkdownPdfSaveOptions? options = null) {
        if (markdown == null) {
            throw new ArgumentNullException(nameof(markdown));
        }

        options ??= new MarkdownPdfSaveOptions();
        MarkdownDoc document = MarkdownReader.Parse(markdown, ResolveReaderOptions(options));
        return document.ToPdfDocument(options);
    }

    /// <summary>
    /// Converts Markdown text to a PDF document and returns conversion diagnostics with it.
    /// </summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this string markdown, MarkdownPdfSaveOptions? options = null) {
        if (markdown == null) {
            throw new ArgumentNullException(nameof(markdown));
        }

        options ??= new MarkdownPdfSaveOptions();
        PdfCore.PdfDocument pdf = markdown.ToPdfDocument(options);
        return new PdfCore.PdfDocumentConversionResult(pdf, options.ConversionReport);
    }

    /// <summary>
    /// Converts a Markdown file to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDocument ToPdfDocumentFromMarkdownFile(this string path, MarkdownPdfSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("Markdown file path cannot be empty.", nameof(path));
        }

        options ??= new MarkdownPdfSaveOptions();
        string fullPath = Path.GetFullPath(path);
        string markdown = File.ReadAllText(fullPath, Encoding.UTF8);
        return MarkdownPdfConverter.ConvertFileMarkdown(markdown, fullPath, options);
    }

    /// <summary>
    /// Converts a Markdown file to a PDF document and returns conversion diagnostics with it.
    /// </summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResultFromMarkdownFile(this string path, MarkdownPdfSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("Markdown file path cannot be empty.", nameof(path));
        }

        options ??= new MarkdownPdfSaveOptions();
        PdfCore.PdfDocument pdf = path.ToPdfDocumentFromMarkdownFile(options);
        return new PdfCore.PdfDocumentConversionResult(pdf, options.ConversionReport);
    }

    /// <summary>
    /// Converts a Markdown document model to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDocument ToPdfDocument(this MarkdownDoc document, MarkdownPdfSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new MarkdownPdfSaveOptions();
        options.ResetExportState();

        PdfCore.PdfOptions pdfOptions = options.PdfOptions?.Clone() ?? new PdfCore.PdfOptions();
        pdfOptions.ReportDiagnosticsTo(options.ConversionReport, "OfficeIMO.Markdown.Pdf");

        if (!string.IsNullOrWhiteSpace(options.FontFamily)) {
            pdfOptions.TryUseOfficeFontFamily(options.FontFamily);
        } else if (options.PdfOptions == null) {
            pdfOptions.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: false);
        }

        if (options.PdfOptions == null) {
            pdfOptions.TryRegisterDefaultDocumentMonospaceFontFallback();
        }

        if (options.CreateOutlineFromHeadings) {
            pdfOptions.CreateOutlineFromHeadings = true;
        }

        MarkdownPdfVisualTheme visualTheme = ResolveVisualTheme(document, options);
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

        options ??= new MarkdownPdfSaveOptions();
        PdfCore.PdfDocument pdf = document.ToPdfDocument(options);
        return new PdfCore.PdfDocumentConversionResult(pdf, options.ConversionReport);
    }

    private static void ApplyMarkdownDefaultFont(PdfCore.PdfDocument pdf, MarkdownPdfSaveOptions options) {
        if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(options.FontFamily, out PdfCore.PdfStandardFont font)) {
            pdf.DefaultTextStyle(style => style.Font(PdfCore.PdfStandardFontMapper.GetFontFamily(font)));
        }
    }

    private static MarkdownReaderOptions ResolveReaderOptions(MarkdownPdfSaveOptions options) {
        if (options.ReaderOptions != null) {
            return options.ReaderOptions;
        }

        MarkdownReaderOptions readerOptions = MarkdownReaderOptions.CreateOfficeIMOProfile();
        readerOptions.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            "OfficeIMO Markdown PDF visual fences",
            new[] { "chart", "ix-chart", "mermaid", "network", "visnetwork", "ix-network", "dataview", "ix-dataview" },
            static context => new SemanticFencedBlock(ResolveVisualFenceSemanticKind(context.Language), context.InfoString, context.Content, context.Caption)));
        return readerOptions;
    }

    private static string ResolveVisualFenceSemanticKind(string? language) {
        switch ((language ?? string.Empty).Trim().ToLowerInvariant()) {
            case "chart":
            case "ix-chart":
                return MarkdownSemanticKinds.Chart;
            case "mermaid":
                return MarkdownSemanticKinds.Mermaid;
            case "network":
            case "visnetwork":
            case "ix-network":
                return MarkdownSemanticKinds.Network;
            case "dataview":
            case "ix-dataview":
                return MarkdownSemanticKinds.DataView;
            default:
                return MarkdownSemanticKinds.Custom;
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
    /// Converts Markdown text to PDF bytes.
    /// </summary>
    public static byte[] SaveAsPdf(this string markdown, MarkdownPdfSaveOptions? options = null) {
        return markdown.ToPdfDocument(options).ToBytes();
    }

    /// <summary>
    /// Saves Markdown text as a PDF file.
    /// </summary>
    public static void SaveAsPdf(this string markdown, string path, MarkdownPdfSaveOptions? options = null) {
        markdown.ToPdfDocument(options).Save(path);
    }

    /// <summary>
    /// Attempts to save Markdown text as a PDF file and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this string markdown, string path, MarkdownPdfSaveOptions? options = null) {
        try {
            return markdown.ToPdfDocument(options).TrySave(path);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>
    /// Writes Markdown text as PDF to a stream.
    /// </summary>
    public static void SaveAsPdf(this string markdown, Stream stream, MarkdownPdfSaveOptions? options = null) {
        markdown.ToPdfDocument(options).Save(stream);
    }

    /// <summary>
    /// Attempts to write Markdown text as PDF to a stream and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this string markdown, Stream stream, MarkdownPdfSaveOptions? options = null) {
        try {
            return markdown.ToPdfDocument(options).TrySave(stream);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    /// <summary>
    /// Converts a Markdown document model to PDF bytes.
    /// </summary>
    public static byte[] SaveAsPdf(this MarkdownDoc document, MarkdownPdfSaveOptions? options = null) {
        return document.ToPdfDocument(options).ToBytes();
    }

    /// <summary>
    /// Saves a Markdown document model as a PDF file.
    /// </summary>
    public static void SaveAsPdf(this MarkdownDoc document, string path, MarkdownPdfSaveOptions? options = null) {
        document.ToPdfDocument(options).Save(path);
    }

    /// <summary>
    /// Attempts to save a Markdown document model as a PDF file and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this MarkdownDoc document, string path, MarkdownPdfSaveOptions? options = null) {
        try {
            return document.ToPdfDocument(options).TrySave(path);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>
    /// Writes a Markdown document model as PDF to a stream.
    /// </summary>
    public static void SaveAsPdf(this MarkdownDoc document, Stream stream, MarkdownPdfSaveOptions? options = null) {
        document.ToPdfDocument(options).Save(stream);
    }

    /// <summary>
    /// Attempts to write a Markdown document model as PDF to a stream and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this MarkdownDoc document, Stream stream, MarkdownPdfSaveOptions? options = null) {
        try {
            return document.ToPdfDocument(options).TrySave(stream);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    private static MarkdownPdfVisualTheme ResolveVisualTheme(MarkdownDoc document, MarkdownPdfSaveOptions options) {
        MarkdownPdfVisualTheme? explicitTheme = options.VisualTheme;
        if (explicitTheme != null) {
            return explicitTheme;
        }

        MarkdownVisualTheme? sharedTheme = options.ThemeSnapshot;
        if (sharedTheme != null) {
            return MarkdownPdfVisualTheme.FromMarkdownTheme(sharedTheme);
        }

        if (options.UseFrontMatterVisualTheme && document.DocumentHeader != null) {
            string? frontMatterPdfTheme = GetFrontMatterMetadata(document.DocumentHeader, "pdfTheme")
                ?? GetFrontMatterMetadata(document.DocumentHeader, "pdf-theme");
            if (frontMatterPdfTheme != null) {
                if (TryResolveSharedOrPdfTheme(frontMatterPdfTheme, sharedFirst: false, out MarkdownPdfVisualTheme? theme)) {
                    return theme!;
                }

                AddWarning(options, "UnsupportedVisualTheme", frontMatterPdfTheme, "The requested Markdown PDF visual theme is not recognized; the configured fallback visual profile is used.");
            }

            string? frontMatterTheme = GetFrontMatterMetadata(document.DocumentHeader, "theme")
                ?? GetFrontMatterMetadata(document.DocumentHeader, "visualTheme")
                ?? GetFrontMatterMetadata(document.DocumentHeader, "visual-theme");
            if (frontMatterTheme != null) {
                if (TryResolveSharedOrPdfTheme(frontMatterTheme, sharedFirst: true, out MarkdownPdfVisualTheme? theme)) {
                    return theme!;
                }

                AddWarning(options, "UnsupportedVisualTheme", frontMatterTheme, "The requested Markdown visual theme is not recognized; the configured fallback visual profile is used.");
            }
        }

        MarkdownVisualTheme? defaultTheme = MarkdownVisualTheme.ResolveOrDefault(null, options.ApplyWordLikeTheme);
        return defaultTheme != null
            ? MarkdownPdfVisualTheme.FromMarkdownTheme(defaultTheme)
            : MarkdownPdfVisualTheme.Plain();
    }

    private static bool TryResolveSharedOrPdfTheme(string themeName, bool sharedFirst, out MarkdownPdfVisualTheme? theme) {
        theme = null;
        if (sharedFirst && MarkdownVisualTheme.TryCreate(themeName, out MarkdownVisualTheme? markdownTheme)) {
            theme = MarkdownPdfVisualTheme.FromMarkdownTheme(markdownTheme!);
            return true;
        }

        if (MarkdownPdfVisualTheme.TryCreate(themeName, out theme)) {
            return true;
        }

        if (!sharedFirst && MarkdownVisualTheme.TryCreate(themeName, out MarkdownVisualTheme? fallbackMarkdownTheme)) {
            theme = MarkdownPdfVisualTheme.FromMarkdownTheme(fallbackMarkdownTheme!);
            return true;
        }

        return false;
    }

    private static void RenderBlocks(PdfCore.PdfDocument pdf, IEnumerable<IMarkdownBlock> blocks, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme, string? skipFirstHeadingTitle = null) {
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
