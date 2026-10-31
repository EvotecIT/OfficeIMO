using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.PdfTextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {


    internal static PdfCore.PdfDocument ConvertToPdfDocument(MarkdownDocument document, MarkdownPdfExportOptions options) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        PdfCore.PdfCreateOptions pdfOptions = options.PdfOptions?.Clone() ?? new PdfCore.PdfCreateOptions();
        pdfOptions.UseContentStreamCompressionByDefault();
        pdfOptions.ReportDiagnosticsTo(options.Report, "OfficeIMO.Markdown.Pdf");

        ApplyMarkdownTextFallbackOptions(pdfOptions, options, document);
        // The writer embeds theme fallback fonts for standard slots unless the resource policy forbids installed fonts.
        if (!options.ResourcePolicy.AllowSystemFontEmbedding) pdfOptions.EmbedDefaultFonts = false;

        // Caller-supplied PDF options own the page furniture, including fonts that may not cover the footer text.
        if (options.ShowPageNumbers && options.PdfOptions == null) {
            pdfOptions.ShowPageNumbers = true;
        }

        if (options.CreateOutlineFromHeadings) {
            pdfOptions.CreateOutlineFromHeadings = true;
        }

        MarkdownPdfStyle visualTheme = ResolveVisualTheme(document, options);
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.CreateForImportedLayout(pdfOptions);
        PdfCore.PdfTheme? documentTheme = visualTheme.DocumentThemeSnapshot;
        if (documentTheme != null) {
            pdf.Theme(documentTheme);
        }
        ApplyMarkdownDefaultFont(pdf, options);
        visualTheme.ApplyPageDecorations(pdf, pdfOptions);

        IReadOnlyList<IMarkdownBlock> topLevelBlocks = GetPdfTopLevelBlocks(document);
        options.Anchors = new AnchorContext(document, topLevelBlocks, options);
        ApplyMetadata(pdf, document, options);
        string? promotedFrontMatterTitle = GetPromotedFrontMatterTitle(document, options);
        RenderBlocks(pdf, topLevelBlocks.Where(block => !(block is FootnoteDefinitionBlock)), document, options, visualTheme, promotedFrontMatterTitle);
        RenderBlocks(pdf, topLevelBlocks.OfType<FootnoteDefinitionBlock>(), document, options, visualTheme);
        if (topLevelBlocks.Count == 0) {
            pdf.Paragraph(paragraph => paragraph.Text(string.Empty));
        }

        return pdf;
    }

    /// <summary>
    /// Converts a Markdown document model to a PDF document and returns conversion diagnostics with it.
    /// </summary>
    internal static PdfCore.PdfDocumentConversionResult ConvertToPdfWithEvidence(this MarkdownDocument document, MarkdownPdfExportOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        MarkdownPdfExportOptions operation = (options ?? new MarkdownPdfExportOptions()).CloneForConversion();
        operation.CancellationToken = cancellationToken;
        // A README converted with Load(path).SaveAsPdf(...) shows its own relative images: they resolve
        // inside the file's folder, while an assigned policy or base directory is kept as given.
        if (operation.BaseDirectory == null && document.FilePath != null && operation.IncludeImages) {
            operation.BaseDirectory = Path.GetDirectoryName(document.FilePath);
            if (!operation.ResourcePolicyAssigned && operation.RestrictLocalImagesToBaseDirectory &&
                !operation.ResourcePolicy.AllowLocalFileAccess) {
                PdfCore.PdfResourcePolicy policy = operation.ResourcePolicy.Clone();
                policy.AllowLocalFileAccess = true;
                operation.ResourcePolicy = policy;
            }
        }
        PdfCore.PdfDocument pdf = ConvertToPdfDocument(document, operation);
        return new PdfCore.PdfDocumentConversionResult(pdf, operation.Report);
    }

    private static void ApplyMarkdownDefaultFont(PdfCore.PdfDocument pdf, MarkdownPdfExportOptions options) {
        if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(options.FontFamily, out PdfCore.PdfStandardFont font)) {
            pdf.DefaultTextStyle(style => style.Font(PdfCore.PdfStandardFontMapper.GetFontFamily(font)));
        }
    }

    private static IReadOnlyList<IMarkdownBlock> GetPdfTopLevelBlocks(MarkdownDocument document) {
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



















    private static MarkdownPdfStyle ResolveVisualTheme(MarkdownDocument document, MarkdownPdfExportOptions options) {
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

                AddWarning(options, "unsupported-visual-theme", frontMatterPdfTheme, "The requested Markdown PDF visual theme is not recognized; the configured fallback visual profile is used.");
            }

            string? frontMatterTheme = GetFrontMatterMetadata(document.DocumentHeader, "theme")
                ?? GetFrontMatterMetadata(document.DocumentHeader, "visualTheme")
                ?? GetFrontMatterMetadata(document.DocumentHeader, "visual-theme");
            if (frontMatterTheme != null) {
                if (TryResolveTheme(frontMatterTheme, out MarkdownPdfStyle? theme)) {
                    return theme!;
                }

                AddWarning(options, "unsupported-visual-theme", frontMatterTheme, "The requested Markdown visual theme is not recognized; the configured fallback visual profile is used.");
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

    private static void RenderBlocks(PdfCore.PdfDocument pdf, IEnumerable<IMarkdownBlock> blocks, MarkdownDocument document, MarkdownPdfExportOptions options, MarkdownPdfStyle visualTheme, string? skipFirstHeadingTitle = null) {
        bool skippedPromotedHeading = false;
        var materializedBlocks = blocks as IReadOnlyList<IMarkdownBlock> ?? blocks.ToList();
        for (int i = 0; i < materializedBlocks.Count; i++) {
            IMarkdownBlock block = materializedBlocks[i];
            if (!skippedPromotedHeading && skipFirstHeadingTitle != null && block is HeadingBlock heading && heading.Level == 1 && IsSameNormalizedText(heading.Text, skipFirstHeadingTitle)) {
                skippedPromotedHeading = true;
                options.Anchors?.Register(pdf, block);
                continue;
            }

            if (block is HeadingBlock tocTitleHeading &&
                i + 1 < materializedBlocks.Count &&
                materializedBlocks[i + 1] is TocBlock toc &&
                ShouldRenderTocAsPanel(toc) &&
                toc.IncludeTitle &&
                IsSameNormalizedText(tocTitleHeading.Text, toc.Title)) {
                options.Anchors?.Register(pdf, block);
                continue;
            }

            // A lead-in such as "Build and preview:" stays on the page of the code block it introduces.
            if (block is ParagraphBlock leadIn && i + 1 < materializedBlocks.Count &&
                materializedBlocks[i + 1] is CodeBlock or SemanticFencedBlock && EndsWithColon(leadIn.Inlines)) {
                options.Anchors?.Register(pdf, block);
                RenderParagraph(pdf, leadIn.Inlines, options, visualTheme, keepWithNext: true);
                continue;
            }

            RenderBlock(pdf, block, document, options, visualTheme);
        }
    }

    private static void ApplyMetadata(PdfCore.PdfDocument pdf, MarkdownDocument document, MarkdownPdfExportOptions options) {
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

    private static string? GetPromotedFrontMatterTitle(MarkdownDocument document, MarkdownPdfExportOptions options) {
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
