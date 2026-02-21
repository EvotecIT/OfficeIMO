using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Markdown → Word converter powered by OfficeIMO.Markdown.
    /// Maps OMD blocks/inlines onto OfficeIMO.Word APIs (headings, lists, tables, images,
    /// code, quotes, callouts, footnotes, etc.).
    /// </summary>
    internal partial class MarkdownToWordConverter {
        private const double DefaultPageWidthTwips = 12240d;
        private const double DefaultHorizontalMarginTwips = 1440d;
        private const double TwipsPerInch = 1440d;
        private const double PixelsPerInch = 96d;

        private static bool LocalPathAllowed(string path, MarkdownToWordOptions options) {
            if (!options.AllowLocalImages) return false;
            if (options.AllowedImageDirectories.Count == 0) return true;
            try {
                var full = System.IO.Path.GetFullPath(path);
                foreach (var root in options.AllowedImageDirectories) {
                    var rootFull = System.IO.Path.GetFullPath(root.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) + System.IO.Path.DirectorySeparatorChar);
                    if (full.StartsWith(rootFull, System.StringComparison.OrdinalIgnoreCase)) return true;
                }
            } catch { return false; }
            return false;
        }

        private static double? ResolveImageWidthLimitPixels(WordDocument document, MarkdownToWordOptions options) {
            var configuredMax = options.MaxImageWidthPixels;
            if (configuredMax.HasValue && configuredMax.Value <= 0) {
                configuredMax = null;
            }

            if (!options.FitImagesToPageContentWidth) {
                return configuredMax;
            }

            var contentWidth = EstimateContentWidthPixels(document);
            if (contentWidth <= 0) {
                return configuredMax;
            }

            return configuredMax.HasValue
                ? Math.Min(configuredMax.Value, contentWidth)
                : contentWidth;
        }

        private static double EstimateContentWidthPixels(WordDocument document) {
            var section = document.Sections.FirstOrDefault();
            var pageWidthTwips = (double?)section?.PageSettings?.Width?.Value ?? DefaultPageWidthTwips;
            var leftMarginTwips = (double?)section?.Margins?.Left?.Value ?? DefaultHorizontalMarginTwips;
            var rightMarginTwips = (double?)section?.Margins?.Right?.Value ?? DefaultHorizontalMarginTwips;
            var contentTwips = pageWidthTwips - leftMarginTwips - rightMarginTwips;

            if (contentTwips < 1) {
                contentTwips = DefaultPageWidthTwips - (DefaultHorizontalMarginTwips * 2);
            }

            if (contentTwips < 1) {
                return 1;
            }

            return contentTwips * PixelsPerInch / TwipsPerInch;
        }

        private static void ConstrainImageDimensions(ref double? width, ref double? height, double? maxWidthPixels) {
            if (!maxWidthPixels.HasValue || maxWidthPixels.Value <= 0) {
                return;
            }

            if (!width.HasValue || width.Value <= 0) {
                return;
            }

            if (width.Value <= maxWidthPixels.Value) {
                return;
            }

            var scale = maxWidthPixels.Value / width.Value;
            width = maxWidthPixels.Value;

            if (height.HasValue && height.Value > 0) {
                height = Math.Max(1, height.Value * scale);
            }
        }

        public WordDocument Convert(string markdown, MarkdownToWordOptions options) {
            return ConvertAsync(markdown, options).GetAwaiter().GetResult();
        }

        public Task<WordDocument> ConvertAsync(string markdown, MarkdownToWordOptions options, CancellationToken cancellationToken = default) {
            if (markdown == null) {
                throw new ArgumentNullException(nameof(markdown));
            }

            options ??= new MarkdownToWordOptions();

            var document = WordDocument.Create();
            options.ApplyDefaults(document);
            var imageWidthLimitPixels = ResolveImageWidthLimitPixels(document, options);

            // Parse using OfficeIMO.Markdown reader.
            var readerOptions = new Omd.MarkdownReaderOptions { BaseUri = options.BaseUri };
            var omd = Omd.MarkdownReader.Parse(markdown, readerOptions);
            var blocks = omd.Blocks;
            // Build footnote definitions map for this document
            _currentFootnotes = blocks is not null
                ? blocks
                    .OfType<Omd.FootnoteDefinitionBlock>()
                    .GroupBy(f => f.Label)
                    .ToDictionary(g => g.Key, g => g.Last().Text)
                : null;

            foreach (var block in blocks ?? Array.Empty<Omd.IMarkdownBlock>()) {
                cancellationToken.ThrowIfCancellationRequested();
                ProcessBlockOmd(block, document, options, quoteDepth: 0, imageWidthLimitPixels: imageWidthLimitPixels);
            }

            return Task.FromResult(document);
        }

        //

        // New OfficeIMO.Markdown path
        private const int IndentTwipsPerLevel = 720; // 0.5 inch per level

        private static void ProcessBlockOmd(
            Omd.IMarkdownBlock block,
            WordDocument document,
            MarkdownToWordOptions options,
            WordList? currentList = null,
            int listLevel = 0,
            int quoteDepth = 0,
            double? imageWidthLimitPixels = null) {
            switch (block) {
                case Omd.HeadingBlock h:
                    var headingParagraph = document.AddParagraph(h.Text ?? string.Empty);
                    if (quoteDepth > 0) headingParagraph.IndentationBefore = IndentTwipsPerLevel * quoteDepth;
                    headingParagraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(h.Level);
                    break;
                case Omd.ParagraphBlock p:
                    if (currentList == null) {
                        var para = document.AddParagraph(string.Empty);
                        if (quoteDepth > 0) para.IndentationBefore = IndentTwipsPerLevel * quoteDepth;
                        ProcessInlinesOmd(p.Inlines, para, options, document, _currentFootnotes);
                    } else {
                        var li = currentList.AddItem(string.Empty, listLevel);
                        if (quoteDepth > 0) li.IndentationBefore = IndentTwipsPerLevel * quoteDepth;
                        ProcessInlinesOmd(p.Inlines, li, options, document, _currentFootnotes);
                    }
                    break;
                case Omd.ImageBlock img: {
                        var par = document.AddParagraph(string.Empty);
                        var pathOrUrl = img.Path ?? string.Empty;
                        double? w = img.Width;
                        double? h = img.Height;
                        if (System.IO.File.Exists(pathOrUrl)) {
                            if (options.AllowLocalImages && LocalPathAllowed(pathOrUrl, options)) {
                                if (w == null || h == null) {
                                    try { using var image = SixLabors.ImageSharp.Image.Load(pathOrUrl, out _); w ??= image.Width; h ??= image.Height; } catch { /* ignore size probe */ }
                                }

                                ConstrainImageDimensions(ref w, ref h, imageWidthLimitPixels);
                                par.AddImage(pathOrUrl, w, h, description: img.Alt ?? string.Empty);
                            } else {
                                // Not allowed: insert as text/link
                                var t1 = par.AddText(img.Alt ?? System.IO.Path.GetFileName(pathOrUrl));
                                if (!string.IsNullOrEmpty(options.FontFamily)) t1.SetFontFamily(options.FontFamily!);
                            }
                        } else if (System.Uri.TryCreate(pathOrUrl, System.UriKind.Absolute, out var uri)) {
                            if (options.AllowedImageSchemes.Contains(uri.Scheme) &&
                                (options.ImageUrlValidator == null || options.ImageUrlValidator(uri))) {
                                if (options.AllowRemoteImages) {
                                    // This call is synchronous inside OfficeIMO.Word; users can choose to disable remote downloads.
                                    ConstrainImageDimensions(ref w, ref h, imageWidthLimitPixels);
                                    document.AddImageFromUrl(uri.ToString(), w, h).Description = img.Alt ?? string.Empty;
                                } else if (options.FallbackRemoteImagesToHyperlinks) {
                                    par.AddHyperLink(img.Alt ?? uri.ToString(), uri);
                                }
                            } else if (options.FallbackRemoteImagesToHyperlinks) {
                                par.AddHyperLink(img.Alt ?? uri.ToString(), uri);
                            }
                        } else {
                            // Not a file or valid URL → insert as text
                            var t2 = par.AddText(img.Alt ?? pathOrUrl);
                            if (!string.IsNullOrEmpty(options.FontFamily)) t2.SetFontFamily(options.FontFamily!);
                        }
                        if (!string.IsNullOrWhiteSpace(img.Caption)) document.AddParagraph(img.Caption!);
                        break;
                    }
                case Omd.CodeBlock cb:
                    var codeParagraph = document.AddParagraph(string.Empty);
                    if (quoteDepth > 0) codeParagraph.IndentationBefore = IndentTwipsPerLevel * quoteDepth;
                    var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";
                    codeParagraph.AddFormattedText(cb.Content ?? string.Empty).SetFontFamily(monoFont);
                    if (!string.IsNullOrWhiteSpace(cb.Caption)) document.AddParagraph(cb.Caption!);
                    break;
                case Omd.TableBlock tb:
                    // Create table and map cell alignments + inline formatting inside cells
                    var cols = (tb.Headers.Count > 0 ? tb.Headers.Count : (tb.Rows.Count > 0 ? tb.Rows[0].Count : 1));
                    var rows = tb.Rows.Count + (tb.Headers.Count > 0 ? 1 : 0);
                    var wtable = document.AddTable(rows: rows, columns: cols);
                    int r = 0;
                    if (tb.Headers.Count > 0) {
                        for (int c = 0; c < tb.Headers.Count; c++) {
                            var para = wtable.Rows[r].Cells[c].Paragraphs[0];
                            ProcessInlinesOmd(Omd.MarkdownReader.ParseInlineText(tb.Headers[c]), para, options, document, _currentFootnotes);
                            // Apply header alignment if provided
                            if (c < tb.Alignments.Count) ApplyAlignment(tb.Alignments[c], para);
                        }
                        r++;
                    }
                    foreach (var row in tb.Rows) {
                        for (int c = 0; c < row.Count && c < wtable.Rows[r].Cells.Count; c++) {
                            var para = wtable.Rows[r].Cells[c].Paragraphs[0];
                            ProcessInlinesOmd(Omd.MarkdownReader.ParseInlineText(row[c]), para, options, document, _currentFootnotes);
                            if (c < tb.Alignments.Count) ApplyAlignment(tb.Alignments[c], para);
                        }
                        r++;
                    }
                    break;
                case Omd.UnorderedListBlock ul: {
                        var list = document.AddList(WordListStyle.Bulleted);
                        foreach (var item in ul.Items) {
                            int effectiveLevel = listLevel + item.Level;
                            var li = list.AddItem(string.Empty, effectiveLevel);
                            // Task list support
                            if (item.IsTask) li.AddCheckBox(item.Checked);
                            ProcessInlinesOmd(item.Content, li, options, document, _currentFootnotes);

                            // Multi-paragraph list items: keep subsequent paragraphs at the same list level.
                            if (item.AdditionalParagraphs != null && item.AdditionalParagraphs.Count > 0) {
                                foreach (var extra in item.AdditionalParagraphs) {
                                    var li2 = list.AddItem(string.Empty, effectiveLevel);
                                    ProcessInlinesOmd(extra, li2, options, document, _currentFootnotes);
                                }
                            }

                            // Nested blocks inside list items (mixed ordered/unordered lists, code blocks, etc.).
                            if (item.Children != null && item.Children.Count > 0) {
                                foreach (var child in item.Children) {
                                    ProcessBlockOmd(child, document, options, null, effectiveLevel + 1, quoteDepth, imageWidthLimitPixels);
                                }
                            }
                        }
                        break;
                    }
                case Omd.OrderedListBlock ol: {
                        var list = document.AddList(WordListStyle.Numbered);
                        if (ol.Start != 1) list.Numbering.Levels[0].SetStartNumberingValue(ol.Start);
                        foreach (var item in ol.Items) {
                            int effectiveLevel = listLevel + item.Level;
                            var li = list.AddItem(string.Empty, effectiveLevel);
                            if (item.IsTask) li.AddCheckBox(item.Checked);
                            ProcessInlinesOmd(item.Content, li, options, document, _currentFootnotes);

                            if (item.AdditionalParagraphs != null && item.AdditionalParagraphs.Count > 0) {
                                foreach (var extra in item.AdditionalParagraphs) {
                                    var li2 = list.AddItem(string.Empty, effectiveLevel);
                                    ProcessInlinesOmd(extra, li2, options, document, _currentFootnotes);
                                }
                            }

                            if (item.Children != null && item.Children.Count > 0) {
                                foreach (var child in item.Children) {
                                    ProcessBlockOmd(child, document, options, null, effectiveLevel + 1, quoteDepth, imageWidthLimitPixels);
                                }
                            }
                        }
                        break;
                    }
                case Omd.TocBlock:
                    // Skip TOC for Word
                    break;
                case Omd.HtmlCommentBlock comment:
                    document.AddHtmlToBody(comment.Comment);
                    break;
                case Omd.HtmlRawBlock html:
                    document.AddHtmlToBody(html.Html);
                    break;
                case Omd.HorizontalRuleBlock:
                    document.AddHorizontalLine();
                    break;
                case Omd.DefinitionListBlock dl:
                    foreach (var (term, definition) in dl.Items) {
                        if (string.IsNullOrWhiteSpace(term) && string.IsNullOrWhiteSpace(definition)) {
                            continue;
                        }

                        var para = document.AddParagraph(string.Empty);
                        if (quoteDepth > 0) {
                            para.IndentationBefore = IndentTwipsPerLevel * quoteDepth;
                        }

                        bool wroteTerm = false;
                        if (!string.IsNullOrWhiteSpace(term)) {
                            var termInlines = Omd.MarkdownReader.ParseInlineText(term);
                            ProcessInlinesOmd(termInlines, para, options, document, _currentFootnotes);
                            wroteTerm = true;
                        }

                        if (!string.IsNullOrWhiteSpace(definition)) {
                            if (wroteTerm) {
                                var separator = para.AddText(": ");
                                if (!string.IsNullOrEmpty(options.FontFamily)) {
                                    separator.SetFontFamily(options.FontFamily!);
                                }
                            }

                            var definitionInlines = Omd.MarkdownReader.ParseInlineText(definition);
                            ProcessInlinesOmd(definitionInlines, para, options, document, _currentFootnotes);
                        }
                    }
                    break;
                case Omd.QuoteBlock qb:
                    foreach (var child in qb.Children) ProcessBlockOmd(child, document, options, null, 0, quoteDepth + 1, imageWidthLimitPixels);
                    break;
                case Omd.CalloutBlock callout:
                    // Render as a simple bold title followed by body paragraphs
                    var ptitle = document.AddParagraph(string.Empty);
                    ptitle.AddFormattedText(callout.Title, bold: true);
                    var pbody = document.AddParagraph(callout.Body);
                    break;
                case Omd.FootnoteDefinitionBlock:
                    // Definitions are consumed when encountering references; skip emitting as body paragraphs.
                    break;
                default:
                    // Fallback: render markdown text
                    var txt = ((Omd.IMarkdownBlock)block).RenderMarkdown();
                    document.AddParagraph(txt);
                    break;
            }
        }

        // Current footnote definitions map; built per-document on ConvertAsync
        private static IReadOnlyDictionary<string, string>? _currentFootnotes;

        private static void ApplyAlignment(Omd.ColumnAlignment align, WordParagraph para) {
            switch (align) {
                case Omd.ColumnAlignment.Left: para.ParagraphAlignment = JustificationValues.Left; break;
                case Omd.ColumnAlignment.Center: para.ParagraphAlignment = JustificationValues.Center; break;
                case Omd.ColumnAlignment.Right: para.ParagraphAlignment = JustificationValues.Right; break;
            }
        }
    }
}
