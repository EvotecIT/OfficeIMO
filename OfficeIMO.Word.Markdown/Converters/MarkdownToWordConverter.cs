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
        private const int IndentTwipsPerLevel = 720; // 0.5 inch per level
        private const double DefaultPageWidthTwips = 12240d;
        private const double DefaultHorizontalMarginTwips = 1440d;
        private const double TwipsPerInch = 1440d;
        private const double PixelsPerInch = 96d;
        private const double MinimumContentWidthPixels = 1d;
        private static readonly TimeSpan DefaultRemoteImageDownloadTimeout = TimeSpan.FromSeconds(20);
        private static readonly System.Net.Http.HttpClient SharedRemoteImageClient = CreateRemoteImageClient(DefaultRemoteImageDownloadTimeout);

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

        private static double EstimatePageContentWidthPixels(WordDocument document) {
            var section = document.Sections.FirstOrDefault();
            var pageWidthTwips = (double?)section?.PageSettings?.Width?.Value ?? DefaultPageWidthTwips;
            var leftMarginTwips = (double?)section?.Margins?.Left?.Value ?? DefaultHorizontalMarginTwips;
            var rightMarginTwips = (double?)section?.Margins?.Right?.Value ?? DefaultHorizontalMarginTwips;
            var contentTwips = pageWidthTwips - leftMarginTwips - rightMarginTwips;

            if (contentTwips < MinimumContentWidthPixels) {
                contentTwips = DefaultPageWidthTwips - (DefaultHorizontalMarginTwips * 2);
            }

            if (contentTwips < MinimumContentWidthPixels) {
                return MinimumContentWidthPixels;
            }

            return contentTwips * PixelsPerInch / TwipsPerInch;
        }

        private static System.Net.Http.HttpClient CreateRemoteImageClient(TimeSpan timeout) {
            var client = new System.Net.Http.HttpClient();
            client.Timeout = timeout;
            client.DefaultRequestHeaders.UserAgent.ParseAdd("OfficeIMO.Word.Markdown");
            return client;
        }

        private static TimeSpan ResolveRemoteImageTimeout(MarkdownToWordOptions options) {
            if (options.RemoteImageDownloadTimeout <= TimeSpan.Zero) {
                return DefaultRemoteImageDownloadTimeout;
            }

            return options.RemoteImageDownloadTimeout;
        }

        private static byte[] DownloadRemoteImageBytes(Uri uri, MarkdownToWordOptions options) {
            var timeout = ResolveRemoteImageTimeout(options);
            if (timeout == DefaultRemoteImageDownloadTimeout) {
                return SharedRemoteImageClient.GetByteArrayAsync(uri).GetAwaiter().GetResult();
            }

            using var client = CreateRemoteImageClient(timeout);
            return client.GetByteArrayAsync(uri).GetAwaiter().GetResult();
        }

        private static double? ResolveContextWidthLimitPixels(
            MarkdownImageLayoutOptions layout,
            double pageContentWidthPixels,
            int listLevel,
            int quoteDepth) {
            if (layout.FitMode == MarkdownImageFitMode.None || pageContentWidthPixels <= 0) {
                return null;
            }

            if (layout.FitMode == MarkdownImageFitMode.PageContentWidth) {
                return pageContentWidthPixels;
            }

            var levels = Math.Max(0, listLevel) + Math.Max(0, quoteDepth);
            if (levels == 0) {
                return pageContentWidthPixels;
            }

            var indentPixels = levels * (IndentTwipsPerLevel * PixelsPerInch / TwipsPerInch);
            return Math.Max(MinimumContentWidthPixels, pageContentWidthPixels - indentPixels);
        }

        private static bool TryGetImageDimensionsFromFile(string filePath, out double width, out double height) {
            width = 0;
            height = 0;
            try {
                using var image = SixLabors.ImageSharp.Image.Load(filePath, out _);
                width = image.Width;
                height = image.Height;
                return width > 0 && height > 0;
            } catch {
                return false;
            }
        }

        private static bool TryGetImageDimensionsFromBytes(byte[] data, out double width, out double height) {
            width = 0;
            height = 0;
            try {
                using var stream = new System.IO.MemoryStream(data, writable: false);
                using var image = SixLabors.ImageSharp.Image.Load(stream, out _);
                width = image.Width;
                height = image.Height;
                return width > 0 && height > 0;
            } catch {
                return false;
            }
        }

        private static bool NormalizePositiveDimension(double? value, out double normalized) {
            normalized = 0;
            if (!value.HasValue || double.IsNaN(value.Value) || double.IsInfinity(value.Value)) {
                return false;
            }

            if (value.Value <= 0) {
                return false;
            }

            normalized = value.Value;
            return true;
        }

        private static void ResolveImageDimensions(
            MarkdownToWordOptions options,
            string source,
            string context,
            double? requestedWidth,
            double? requestedHeight,
            double? naturalWidth,
            double? naturalHeight,
            double pageContentWidthPixels,
            double? contextWidthLimitPixels,
            out double? finalWidth,
            out double? finalHeight,
            out bool scaledByLayout) {
            var layout = options.ImageLayout ?? new MarkdownImageLayoutOptions();
            finalWidth = null;
            finalHeight = null;
            scaledByLayout = false;

            var hasNaturalWidth = NormalizePositiveDimension(naturalWidth, out var naturalWidthPx);
            var hasNaturalHeight = NormalizePositiveDimension(naturalHeight, out var naturalHeightPx);
            var hasRequestedWidth = NormalizePositiveDimension(requestedWidth, out var requestedWidthPx);
            var hasRequestedHeight = NormalizePositiveDimension(requestedHeight, out var requestedHeightPx);

            if (layout.HintPrecedence == MarkdownImageHintPrecedence.LayoutThenMarkdown) {
                if (hasNaturalWidth) {
                    finalWidth = naturalWidthPx;
                }
                if (hasNaturalHeight) {
                    finalHeight = naturalHeightPx;
                }

                if (hasRequestedWidth) {
                    finalWidth = requestedWidthPx;
                }
                if (hasRequestedHeight) {
                    finalHeight = requestedHeightPx;
                }
            } else {
                if (hasRequestedWidth) {
                    finalWidth = requestedWidthPx;
                } else if (hasNaturalWidth) {
                    finalWidth = naturalWidthPx;
                }

                if (hasRequestedHeight) {
                    finalHeight = requestedHeightPx;
                } else if (hasNaturalHeight) {
                    finalHeight = naturalHeightPx;
                }
            }

            if (finalWidth.HasValue && !finalHeight.HasValue && hasNaturalWidth && hasNaturalHeight) {
                finalHeight = naturalHeightPx * (finalWidth.Value / naturalWidthPx);
            } else if (!finalWidth.HasValue && finalHeight.HasValue && hasNaturalWidth && hasNaturalHeight) {
                finalWidth = naturalWidthPx * (finalHeight.Value / naturalHeightPx);
            }

            double? effectiveMaxWidth = null;
            double? effectiveMaxHeight = null;

            if (NormalizePositiveDimension(layout.MaxWidthPixels, out var maxWidth)) {
                effectiveMaxWidth = maxWidth;
            }
            if (NormalizePositiveDimension(layout.MaxHeightPixels, out var maxHeight)) {
                effectiveMaxHeight = maxHeight;
            }
            if (NormalizePositiveDimension(layout.MaxWidthPercentOfContent, out var maxWidthPercent)) {
                var widthBaseline = NormalizePositiveDimension(contextWidthLimitPixels, out var contextWidth)
                    ? contextWidth
                    : (pageContentWidthPixels > 0 ? pageContentWidthPixels : 0);
                if (widthBaseline > 0) {
                    var percentCapWidth = widthBaseline * (maxWidthPercent / 100d);
                    if (percentCapWidth > 0) {
                        effectiveMaxWidth = effectiveMaxWidth.HasValue
                            ? Math.Min(effectiveMaxWidth.Value, percentCapWidth)
                            : percentCapWidth;
                    }
                }
            }
            if (NormalizePositiveDimension(contextWidthLimitPixels, out var contextMaxWidth)) {
                effectiveMaxWidth = effectiveMaxWidth.HasValue ? Math.Min(effectiveMaxWidth.Value, contextMaxWidth) : contextMaxWidth;
            }

            if (!layout.AllowUpscale) {
                if (hasNaturalWidth) {
                    effectiveMaxWidth = effectiveMaxWidth.HasValue ? Math.Min(effectiveMaxWidth.Value, naturalWidthPx) : naturalWidthPx;
                }
                if (hasNaturalHeight) {
                    effectiveMaxHeight = effectiveMaxHeight.HasValue ? Math.Min(effectiveMaxHeight.Value, naturalHeightPx) : naturalHeightPx;
                }
            }

            if (finalWidth.HasValue && finalHeight.HasValue) {
                var scale = 1d;
                if (effectiveMaxWidth.HasValue && finalWidth.Value > effectiveMaxWidth.Value) {
                    scale = Math.Min(scale, effectiveMaxWidth.Value / finalWidth.Value);
                }
                if (effectiveMaxHeight.HasValue && finalHeight.Value > effectiveMaxHeight.Value) {
                    scale = Math.Min(scale, effectiveMaxHeight.Value / finalHeight.Value);
                }
                if (scale < 1d) {
                    finalWidth *= scale;
                    finalHeight *= scale;
                    scaledByLayout = true;
                }
            } else {
                if (finalWidth.HasValue && effectiveMaxWidth.HasValue && finalWidth.Value > effectiveMaxWidth.Value) {
                    finalWidth = effectiveMaxWidth.Value;
                    scaledByLayout = true;
                }
                if (finalHeight.HasValue && effectiveMaxHeight.HasValue && finalHeight.Value > effectiveMaxHeight.Value) {
                    finalHeight = effectiveMaxHeight.Value;
                    scaledByLayout = true;
                }
            }

            if (finalWidth.HasValue && finalWidth.Value <= 0) {
                finalWidth = null;
            }
            if (finalHeight.HasValue && finalHeight.Value <= 0) {
                finalHeight = null;
            }

            if (options.OnImageLayoutDiagnostic != null) {
                options.OnImageLayoutDiagnostic(new MarkdownImageLayoutDiagnostic {
                    Source = source,
                    Context = context,
                    RequestedWidthPixels = hasRequestedWidth ? requestedWidthPx : null,
                    RequestedHeightPixels = hasRequestedHeight ? requestedHeightPx : null,
                    NaturalWidthPixels = hasNaturalWidth ? naturalWidthPx : null,
                    NaturalHeightPixels = hasNaturalHeight ? naturalHeightPx : null,
                    EffectiveMaxWidthPixels = effectiveMaxWidth,
                    EffectiveMaxHeightPixels = effectiveMaxHeight,
                    FinalWidthPixels = finalWidth,
                    FinalHeightPixels = finalHeight,
                    ScaledByLayout = scaledByLayout
                });
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
            var pageContentWidthPixels = EstimatePageContentWidthPixels(document);

            // Parse using OfficeIMO.Markdown reader.
            var readerOptions = new Omd.MarkdownReaderOptions {
                BaseUri = options.BaseUri,
                DefinitionLists = !options.PreferNarrativeSingleLineDefinitions
            };
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
                ProcessBlockOmd(block, document, options, quoteDepth: 0, pageContentWidthPixels: pageContentWidthPixels);
            }

            return Task.FromResult(document);
        }

        //
        private static void ProcessBlockOmd(
            Omd.IMarkdownBlock block,
            WordDocument document,
            MarkdownToWordOptions options,
            WordList? currentList = null,
            int listLevel = 0,
            int quoteDepth = 0,
            double pageContentWidthPixels = 0) {
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
                        var contextWidthLimit = ResolveContextWidthLimitPixels(options.ImageLayout, pageContentWidthPixels, listLevel, quoteDepth);

                        if (System.IO.File.Exists(pathOrUrl)) {
                            if (options.AllowLocalImages && LocalPathAllowed(pathOrUrl, options)) {
                                double? naturalW = null;
                                double? naturalH = null;
                                if (TryGetImageDimensionsFromFile(pathOrUrl, out var fileW, out var fileH)) {
                                    naturalW = fileW;
                                    naturalH = fileH;
                                }

                                ResolveImageDimensions(
                                    options,
                                    source: pathOrUrl,
                                    context: "block-local",
                                    requestedWidth: img.Width,
                                    requestedHeight: img.Height,
                                    naturalWidth: naturalW,
                                    naturalHeight: naturalH,
                                    pageContentWidthPixels: pageContentWidthPixels,
                                    contextWidthLimitPixels: contextWidthLimit,
                                    out var finalW,
                                    out var finalH,
                                    out _);

                                par.AddImage(pathOrUrl, finalW, finalH, description: img.Alt ?? string.Empty);
                            } else {
                                // Not allowed: insert as text/link
                                var t1 = par.AddText(img.Alt ?? System.IO.Path.GetFileName(pathOrUrl));
                                if (!string.IsNullOrEmpty(options.FontFamily)) t1.SetFontFamily(options.FontFamily!);
                            }
                        } else if (System.Uri.TryCreate(pathOrUrl, System.UriKind.Absolute, out var uri)) {
                            if (options.AllowedImageSchemes.Contains(uri.Scheme) &&
                                (options.ImageUrlValidator == null || options.ImageUrlValidator(uri))) {
                                if (options.AllowRemoteImages) {
                                    try {
                                        var bytes = DownloadRemoteImageBytes(uri, options);
                                        var fileName = System.IO.Path.GetFileName(uri.LocalPath);
                                        if (string.IsNullOrWhiteSpace(fileName)) {
                                            fileName = "image";
                                        }

                                        double? naturalW = null;
                                        double? naturalH = null;
                                        if (TryGetImageDimensionsFromBytes(bytes, out var remoteW, out var remoteH)) {
                                            naturalW = remoteW;
                                            naturalH = remoteH;
                                        }

                                        ResolveImageDimensions(
                                            options,
                                            source: uri.ToString(),
                                            context: "block-remote",
                                            requestedWidth: img.Width,
                                            requestedHeight: img.Height,
                                            naturalWidth: naturalW,
                                            naturalHeight: naturalH,
                                            pageContentWidthPixels: pageContentWidthPixels,
                                            contextWidthLimitPixels: contextWidthLimit,
                                            out var finalW,
                                            out var finalH,
                                            out _);

                                        using var stream = new System.IO.MemoryStream(bytes, writable: false);
                                        par.AddImage(stream, fileName, finalW, finalH, description: img.Alt ?? string.Empty);
                                    } catch (Exception ex) {
                                        options.OnWarning?.Invoke($"Remote image '{uri}' could not be downloaded. {ex.Message}");
                                        if (options.FallbackRemoteImagesToHyperlinks) {
                                            par.AddHyperLink(img.Alt ?? uri.ToString(), uri);
                                        }
                                    }
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
                                    ProcessBlockOmd(child, document, options, null, effectiveLevel + 1, quoteDepth, pageContentWidthPixels);
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
                                    ProcessBlockOmd(child, document, options, null, effectiveLevel + 1, quoteDepth, pageContentWidthPixels);
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
                    foreach (var child in qb.Children) ProcessBlockOmd(child, document, options, null, 0, quoteDepth + 1, pageContentWidthPixels);
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
