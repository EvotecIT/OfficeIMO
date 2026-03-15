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

        private static System.Net.Http.HttpClient CreateRemoteImageClient(TimeSpan timeout, bool bypassProxy = false) {
            System.Net.Http.HttpClient client;
            if (bypassProxy) {
                var handler = new System.Net.Http.HttpClientHandler {
                    Proxy = null,
                    UseProxy = false
                };
                client = new System.Net.Http.HttpClient(handler, disposeHandler: true);
            } else {
                client = new System.Net.Http.HttpClient();
            }

            client.Timeout = timeout;
            client.DefaultRequestHeaders.UserAgent.ParseAdd("OfficeIMO.Word.Markdown");
            return client;
        }

        private static bool IsLoopbackImageUri(Uri uri) {
            if (uri == null) {
                return false;
            }

            if (uri.IsLoopback) {
                return true;
            }

            return string.Equals(uri.Host, "localhost", StringComparison.OrdinalIgnoreCase);
        }

        private static string? ResolveDefaultFontFamily(MarkdownToWordOptions options) {
            if (options == null) {
                return null;
            }

            return FontResolver.Resolve(options.FontFamily) ?? options.FontFamily;
        }

        private static TimeSpan ResolveRemoteImageTimeout(MarkdownToWordOptions options) {
            if (options.RemoteImageDownloadTimeout <= TimeSpan.Zero) {
                return DefaultRemoteImageDownloadTimeout;
            }

            return options.RemoteImageDownloadTimeout;
        }

        private static Omd.MarkdownReaderOptions CreateEffectiveReaderOptions(MarkdownToWordOptions options) {
            var source = options.ReaderOptions;
            if (source == null) {
                return new Omd.MarkdownReaderOptions {
                    BaseUri = options.BaseUri,
                    PreferNarrativeSingleLineDefinitions = options.PreferNarrativeSingleLineDefinitions
                };
            }

            var effective = new Omd.MarkdownReaderOptions {
                FrontMatter = source.FrontMatter,
                Callouts = source.Callouts,
                Headings = source.Headings,
                FencedCode = source.FencedCode,
                IndentedCodeBlocks = source.IndentedCodeBlocks,
                Images = source.Images,
                UnorderedLists = source.UnorderedLists,
                TaskLists = source.TaskLists,
                OrderedLists = source.OrderedLists,
                Tables = source.Tables,
                DefinitionLists = source.DefinitionLists,
                TocPlaceholders = source.TocPlaceholders,
                Footnotes = source.Footnotes,
                PreferNarrativeSingleLineDefinitions = source.PreferNarrativeSingleLineDefinitions,
                HtmlBlocks = source.HtmlBlocks,
                Paragraphs = source.Paragraphs,
                AutolinkUrls = source.AutolinkUrls,
                AutolinkWwwUrls = source.AutolinkWwwUrls,
                AutolinkWwwScheme = source.AutolinkWwwScheme,
                AutolinkEmails = source.AutolinkEmails,
                BackslashHardBreaks = source.BackslashHardBreaks,
                InlineHtml = source.InlineHtml,
                BaseUri = source.BaseUri,
                DisallowScriptUrls = source.DisallowScriptUrls,
                DisallowFileUrls = source.DisallowFileUrls,
                AllowMailtoUrls = source.AllowMailtoUrls,
                AllowDataUrls = source.AllowDataUrls,
                AllowProtocolRelativeUrls = source.AllowProtocolRelativeUrls,
                RestrictUrlSchemes = source.RestrictUrlSchemes,
                AllowedUrlSchemes = source.AllowedUrlSchemes?.ToArray() ?? Array.Empty<string>(),
                InputNormalization = new Omd.MarkdownInputNormalizationOptions {
                    NormalizeSoftWrappedStrongSpans = source.InputNormalization?.NormalizeSoftWrappedStrongSpans ?? false,
                    NormalizeInlineCodeSpanLineBreaks = source.InputNormalization?.NormalizeInlineCodeSpanLineBreaks ?? false,
                    NormalizeEscapedInlineCodeSpans = source.InputNormalization?.NormalizeEscapedInlineCodeSpans ?? false,
                    NormalizeTightStrongBoundaries = source.InputNormalization?.NormalizeTightStrongBoundaries ?? false,
                    NormalizeTightArrowStrongBoundaries = source.InputNormalization?.NormalizeTightArrowStrongBoundaries ?? false,
                    NormalizeBrokenStrongArrowLabels = source.InputNormalization?.NormalizeBrokenStrongArrowLabels ?? false,
                    NormalizeWrappedSignalFlowStrongRuns = source.InputNormalization?.NormalizeWrappedSignalFlowStrongRuns ?? false,
                    NormalizeCollapsedMetricChains = source.InputNormalization?.NormalizeCollapsedMetricChains ?? false,
                    NormalizeHostLabelBulletArtifacts = source.InputNormalization?.NormalizeHostLabelBulletArtifacts ?? false,
                    NormalizeTightColonSpacing = source.InputNormalization?.NormalizeTightColonSpacing ?? false,
                    NormalizeHeadingListBoundaries = source.InputNormalization?.NormalizeHeadingListBoundaries ?? false,
                    NormalizeCompactStrongLabelListBoundaries = source.InputNormalization?.NormalizeCompactStrongLabelListBoundaries ?? false,
                    NormalizeCompactHeadingBoundaries = source.InputNormalization?.NormalizeCompactHeadingBoundaries ?? false,
                    NormalizeStandaloneHashHeadingSeparators = source.InputNormalization?.NormalizeStandaloneHashHeadingSeparators ?? false,
                    NormalizeBrokenTwoLineStrongLeadIns = source.InputNormalization?.NormalizeBrokenTwoLineStrongLeadIns ?? false,
                    NormalizeColonListBoundaries = source.InputNormalization?.NormalizeColonListBoundaries ?? false,
                    NormalizeCompactFenceBodyBoundaries = source.InputNormalization?.NormalizeCompactFenceBodyBoundaries ?? false,
                    NormalizeLooseStrongDelimiters = source.InputNormalization?.NormalizeLooseStrongDelimiters ?? false,
                    NormalizeOrderedListMarkerSpacing = source.InputNormalization?.NormalizeOrderedListMarkerSpacing ?? false,
                    NormalizeOrderedListParenMarkers = source.InputNormalization?.NormalizeOrderedListParenMarkers ?? false,
                    NormalizeOrderedListCaretArtifacts = source.InputNormalization?.NormalizeOrderedListCaretArtifacts ?? false,
                    NormalizeTightParentheticalSpacing = source.InputNormalization?.NormalizeTightParentheticalSpacing ?? false,
                    NormalizeNestedStrongDelimiters = source.InputNormalization?.NormalizeNestedStrongDelimiters ?? false,
                    NormalizeDanglingTrailingStrongListClosers = source.InputNormalization?.NormalizeDanglingTrailingStrongListClosers ?? false,
                    NormalizeMetricValueStrongRuns = source.InputNormalization?.NormalizeMetricValueStrongRuns ?? false
                }
            };

            if (string.IsNullOrWhiteSpace(effective.BaseUri) && !string.IsNullOrWhiteSpace(options.BaseUri)) {
                effective.BaseUri = options.BaseUri;
            }

            if (!effective.PreferNarrativeSingleLineDefinitions && options.PreferNarrativeSingleLineDefinitions) {
                effective.PreferNarrativeSingleLineDefinitions = true;
            }

            CopyBlockParserExtensions(source, effective);
            CopyFencedBlockExtensions(source, effective);
            CopyDocumentTransforms(source, effective);
            return effective;
        }

        private static void CopyFencedBlockExtensions(Omd.MarkdownReaderOptions source, Omd.MarkdownReaderOptions target) {
            if (source.FencedBlockExtensions == null || source.FencedBlockExtensions.Count == 0) {
                return;
            }

            for (int i = 0; i < source.FencedBlockExtensions.Count; i++) {
                var extension = source.FencedBlockExtensions[i];
                if (extension != null) {
                    target.FencedBlockExtensions.Add(extension);
                }
            }
        }

        private static void CopyBlockParserExtensions(Omd.MarkdownReaderOptions source, Omd.MarkdownReaderOptions target) {
            target.BlockParserExtensions.Clear();
            if (source.BlockParserExtensions == null || source.BlockParserExtensions.Count == 0) {
                return;
            }

            for (int i = 0; i < source.BlockParserExtensions.Count; i++) {
                var extension = source.BlockParserExtensions[i];
                if (extension != null) {
                    target.BlockParserExtensions.Add(extension);
                }
            }
        }

        private static void CopyDocumentTransforms(Omd.MarkdownReaderOptions source, Omd.MarkdownReaderOptions target) {
            if (source.DocumentTransforms == null || source.DocumentTransforms.Count == 0) {
                return;
            }

            for (var i = 0; i < source.DocumentTransforms.Count; i++) {
                var transform = source.DocumentTransforms[i];
                if (transform != null) {
                    target.DocumentTransforms.Add(transform);
                }
            }
        }

        private static byte[] DownloadRemoteImageBytes(Uri uri, MarkdownToWordOptions options) {
            var timeout = ResolveRemoteImageTimeout(options);
            // Fresh clients avoid stale loopback/proxy behavior on older framework handlers.
            using var client = CreateRemoteImageClient(timeout, bypassProxy: IsLoopbackImageUri(uri));
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

        public WordDocument Convert(Omd.MarkdownDoc markdown, MarkdownToWordOptions options) {
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
            var readerOptions = CreateEffectiveReaderOptions(options);
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

        public Task<WordDocument> ConvertAsync(Omd.MarkdownDoc markdown, MarkdownToWordOptions options, CancellationToken cancellationToken = default) {
            if (markdown == null) {
                throw new ArgumentNullException(nameof(markdown));
            }

            options ??= new MarkdownToWordOptions();

            var document = WordDocument.Create();
            options.ApplyDefaults(document);
            var pageContentWidthPixels = EstimatePageContentWidthPixels(document);
            var blocks = markdown.Blocks ?? Array.Empty<Omd.IMarkdownBlock>();

            _currentFootnotes = blocks
                .OfType<Omd.FootnoteDefinitionBlock>()
                .GroupBy(f => f.Label)
                .ToDictionary(g => g.Key, g => g.Last().Text);

            var host = new DocumentWordBlockRenderHost(document);
            foreach (var block in blocks) {
                cancellationToken.ThrowIfCancellationRequested();
                if (block == null) {
                    continue;
                }

                RenderSharedBlockOmd(
                    block,
                    host,
                    options,
                    document,
                    currentList: null,
                    listLevel: 0,
                    quoteDepth: 0,
                    pageContentWidthPixels: pageContentWidthPixels,
                    alignment: Omd.ColumnAlignment.None);
            }

            return Task.FromResult(document);
        }

        private interface IWordBlockRenderHost {
            WordParagraph CreateParagraph();
            WordList CreateList(WordListStyle style);
            WordTable CreateTable(int rows, int columns);
            bool SupportsHtmlInsertion { get; }
            void InsertHtml(string html);
            bool SupportsHorizontalRule { get; }
            void InsertHorizontalRule();
        }

        private sealed class DocumentWordBlockRenderHost : IWordBlockRenderHost {
            private readonly WordDocument _document;

            public DocumentWordBlockRenderHost(WordDocument document) {
                _document = document ?? throw new ArgumentNullException(nameof(document));
            }

            public WordParagraph CreateParagraph() => _document.AddParagraph(string.Empty);
            public WordList CreateList(WordListStyle style) => _document.AddList(style);
            public WordTable CreateTable(int rows, int columns) => _document.AddTable(rows, columns);
            public bool SupportsHtmlInsertion => true;
            public void InsertHtml(string html) => _document.AddHtmlToBody(html);
            public bool SupportsHorizontalRule => true;
            public void InsertHorizontalRule() => _document.AddHorizontalLine();
        }

        private sealed class TableCellWordBlockRenderHost : IWordBlockRenderHost {
            private readonly WordTableCell _cell;
            private bool _wroteContent;

            public TableCellWordBlockRenderHost(WordTableCell cell) {
                _cell = cell ?? throw new ArgumentNullException(nameof(cell));
            }

            public WordParagraph CreateParagraph() {
                if (!_wroteContent) {
                    var existing = _cell.Paragraphs.FirstOrDefault();
                    if (existing != null) {
                        _wroteContent = true;
                        return existing;
                    }
                }

                _wroteContent = true;
                return _cell.AddParagraph();
            }

            public WordList CreateList(WordListStyle style) {
                _wroteContent = true;
                return _cell.AddList(style);
            }

            public WordTable CreateTable(int rows, int columns) {
                _wroteContent = true;
                return _cell.AddTable(rows, columns);
            }

            public bool SupportsHtmlInsertion => false;
            public void InsertHtml(string html) { }
            public bool SupportsHorizontalRule => false;
            public void InsertHorizontalRule() { }
        }

        private static void ApplyBlockParagraphFormatting(WordParagraph paragraph, int quoteDepth, Omd.ColumnAlignment alignment) {
            if (quoteDepth > 0) {
                paragraph.IndentationBefore = IndentTwipsPerLevel * quoteDepth;
            }

            ApplyAlignment(alignment, paragraph);
        }

        private static void ProcessTableCellBlocksOmd(
            Omd.TableCell? tableCell,
            WordTableCell wordCell,
            MarkdownToWordOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.ColumnAlignment alignment) {
            if (tableCell == null || tableCell.Blocks.Count == 0) {
                return;
            }

            var host = new TableCellWordBlockRenderHost(wordCell);
            RenderSharedBlocksOmd(tableCell.Blocks, host, options, document, quoteDepth: quoteDepth, pageContentWidthPixels: pageContentWidthPixels, alignment: alignment);
        }

        private static void RenderSharedBlocksOmd(
            IEnumerable<Omd.IMarkdownBlock> blocks,
            IWordBlockRenderHost host,
            MarkdownToWordOptions options,
            WordDocument document,
            int listLevel = 0,
            int quoteDepth = 0,
            double pageContentWidthPixels = 0,
            Omd.ColumnAlignment alignment = Omd.ColumnAlignment.None) {
            if (blocks == null) {
                return;
            }

            foreach (var block in blocks) {
                if (block == null) {
                    continue;
                }

                RenderSharedBlockOmd(block, host, options, document, currentList: null, listLevel, quoteDepth, pageContentWidthPixels, alignment);
            }
        }

        private static void RenderSharedDefinitionListEntryOmd(
            Omd.DefinitionListEntry entry,
            IWordBlockRenderHost host,
            MarkdownToWordOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.ColumnAlignment alignment) {
            if (entry == null) {
                return;
            }

            bool hasTerm = !string.IsNullOrWhiteSpace(entry.TermMarkdown);
            int nextDefinitionBlockIndex = 0;
            WordParagraph? leadParagraph = null;

            if (hasTerm || (entry.DefinitionBlocks.Count > 0 && entry.DefinitionBlocks[0] is Omd.ParagraphBlock)) {
                leadParagraph = host.CreateParagraph();
                ApplyBlockParagraphFormatting(leadParagraph, quoteDepth, alignment);
            }

            if (hasTerm && leadParagraph != null) {
                ProcessInlinesOmd(entry.Term, leadParagraph, options, document, _currentFootnotes);
            }

            if (entry.DefinitionBlocks.Count > 0 && entry.DefinitionBlocks[0] is Omd.ParagraphBlock firstParagraph) {
                if (leadParagraph == null) {
                    leadParagraph = host.CreateParagraph();
                    ApplyBlockParagraphFormatting(leadParagraph, quoteDepth, alignment);
                }

                if (hasTerm) {
                    var separator = leadParagraph.AddText(": ");
                    var defaultFont = ResolveDefaultFontFamily(options);
                    if (!string.IsNullOrEmpty(defaultFont)) {
                        separator.SetFontFamily(defaultFont!);
                    }
                }

                ProcessInlinesOmd(firstParagraph.Inlines, leadParagraph, options, document, _currentFootnotes);
                nextDefinitionBlockIndex = 1;
            }

            if (entry.DefinitionBlocks.Count == 0 && hasTerm && leadParagraph == null) {
                leadParagraph = host.CreateParagraph();
                ApplyBlockParagraphFormatting(leadParagraph, quoteDepth, alignment);
                ProcessInlinesOmd(entry.Term, leadParagraph, options, document, _currentFootnotes);
            }

            for (int i = nextDefinitionBlockIndex; i < entry.DefinitionBlocks.Count; i++) {
                RenderSharedBlockOmd(
                    entry.DefinitionBlocks[i],
                    host,
                    options,
                    document,
                    currentList: null,
                    listLevel: 0,
                    quoteDepth: quoteDepth,
                    pageContentWidthPixels: pageContentWidthPixels,
                    alignment: alignment);
            }
        }

        private static void RenderSharedTableBlockOmd(
            Omd.TableBlock table,
            IWordBlockRenderHost host,
            MarkdownToWordOptions options,
            WordDocument document,
            double pageContentWidthPixels) {
            var headerCells = table.HeaderCells;
            var rowCells = table.RowCells;
            var cols = headerCells.Count > 0
                ? headerCells.Count
                : (rowCells.Count > 0 ? rowCells[0].Count : 1);
            var rows = rowCells.Count + (headerCells.Count > 0 ? 1 : 0);
            var wordTable = host.CreateTable(rows, cols);
            int rowIndex = 0;

            if (headerCells.Count > 0) {
                for (int columnIndex = 0; columnIndex < headerCells.Count; columnIndex++) {
                    var alignment = columnIndex < table.Alignments.Count ? table.Alignments[columnIndex] : Omd.ColumnAlignment.None;
                    var cellHost = new TableCellWordBlockRenderHost(wordTable.Rows[rowIndex].Cells[columnIndex]);
                    RenderSharedBlocksOmd(
                        headerCells[columnIndex].Blocks,
                        cellHost,
                        options,
                        document,
                        quoteDepth: 0,
                        pageContentWidthPixels: pageContentWidthPixels,
                        alignment: alignment);
                }
                rowIndex++;
            }

            for (int sourceRowIndex = 0; sourceRowIndex < rowCells.Count; sourceRowIndex++) {
                var row = rowCells[sourceRowIndex];
                for (int columnIndex = 0; columnIndex < row.Count && columnIndex < wordTable.Rows[rowIndex].Cells.Count; columnIndex++) {
                    var alignment = columnIndex < table.Alignments.Count ? table.Alignments[columnIndex] : Omd.ColumnAlignment.None;
                    var cellHost = new TableCellWordBlockRenderHost(wordTable.Rows[rowIndex].Cells[columnIndex]);
                    RenderSharedBlocksOmd(
                        row[columnIndex].Blocks,
                        cellHost,
                        options,
                        document,
                        quoteDepth: 0,
                        pageContentWidthPixels: pageContentWidthPixels,
                        alignment: alignment);
                }
                rowIndex++;
            }
        }

        private static void RenderSharedCalloutBlockOmd(
            Omd.CalloutBlock callout,
            IWordBlockRenderHost host,
            MarkdownToWordOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.ColumnAlignment alignment) {
            var titleParagraph = host.CreateParagraph();
            ApplyBlockParagraphFormatting(titleParagraph, quoteDepth, alignment);
            if (callout.TitleInlines != null && (callout.TitleInlines.Items?.Count ?? 0) > 0) {
                ProcessInlinesOmd(callout.TitleInlines, titleParagraph, options, document, _currentFootnotes);
                foreach (var run in titleParagraph.GetRuns()) {
                    run.SetBold();
                }
            } else {
                titleParagraph.AddFormattedText(callout.Title, bold: true);
            }

            if (callout.ChildBlocks.Count > 0) {
                RenderSharedBlocksOmd(callout.ChildBlocks, host, options, document, quoteDepth: quoteDepth, pageContentWidthPixels: pageContentWidthPixels, alignment: alignment);
            } else if (!string.IsNullOrWhiteSpace(callout.Body)) {
                var bodyParagraph = host.CreateParagraph();
                ApplyBlockParagraphFormatting(bodyParagraph, quoteDepth, alignment);
                bodyParagraph.AddText(callout.Body);
            }
        }

        private static void RenderSharedBlockOmd(
            Omd.IMarkdownBlock block,
            IWordBlockRenderHost host,
            MarkdownToWordOptions options,
            WordDocument document,
            WordList? currentList = null,
            int listLevel = 0,
            int quoteDepth = 0,
            double pageContentWidthPixels = 0,
            Omd.ColumnAlignment alignment = Omd.ColumnAlignment.None) {
            switch (block) {
                case Omd.HeadingBlock h: {
                        var headingParagraph = host.CreateParagraph();
                        ApplyBlockParagraphFormatting(headingParagraph, quoteDepth, alignment);
                        headingParagraph.SetText(h.Text ?? string.Empty);
                        headingParagraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(h.Level);
                        break;
                    }
                case Omd.ParagraphBlock p:
                    if (currentList == null) {
                        var para = host.CreateParagraph();
                        ApplyBlockParagraphFormatting(para, quoteDepth, alignment);
                        ProcessInlinesOmd(p.Inlines, para, options, document, _currentFootnotes);
                    } else {
                        var li = currentList.AddItem(string.Empty, listLevel);
                        ApplyBlockParagraphFormatting(li, quoteDepth, alignment);
                        ProcessInlinesOmd(p.Inlines, li, options, document, _currentFootnotes);
                    }
                    break;
                case Omd.ImageBlock img: {
                        var par = host.CreateParagraph();
                        ApplyBlockParagraphFormatting(par, quoteDepth, alignment);
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
                                var t1 = par.AddText(img.Alt ?? System.IO.Path.GetFileName(pathOrUrl));
                                var defaultFont = ResolveDefaultFontFamily(options);
                                if (!string.IsNullOrEmpty(defaultFont)) {
                                    t1.SetFontFamily(defaultFont!);
                                }
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
                            var t2 = par.AddText(img.Alt ?? pathOrUrl);
                            var defaultFont = ResolveDefaultFontFamily(options);
                            if (!string.IsNullOrEmpty(defaultFont)) {
                                t2.SetFontFamily(defaultFont!);
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(img.Caption)) {
                            var captionParagraph = host.CreateParagraph();
                            ApplyBlockParagraphFormatting(captionParagraph, quoteDepth, alignment);
                            captionParagraph.AddText(img.Caption!);
                        }
                        break;
                    }
                case Omd.CodeBlock cb: {
                        var codeParagraph = host.CreateParagraph();
                        ApplyBlockParagraphFormatting(codeParagraph, quoteDepth, alignment);
                        var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";
                        codeParagraph.AddFormattedText(cb.Content ?? string.Empty).SetFontFamily(monoFont);
                        if (!string.IsNullOrWhiteSpace(cb.Caption)) {
                            var captionParagraph = host.CreateParagraph();
                            ApplyBlockParagraphFormatting(captionParagraph, quoteDepth, alignment);
                            captionParagraph.AddText(cb.Caption!);
                        }
                        break;
                    }
                case Omd.SemanticFencedBlock semantic: {
                        var semanticParagraph = host.CreateParagraph();
                        ApplyBlockParagraphFormatting(semanticParagraph, quoteDepth, alignment);
                        var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";
                        semanticParagraph.AddFormattedText(semantic.Content ?? string.Empty).SetFontFamily(monoFont);
                        if (!string.IsNullOrWhiteSpace(semantic.Caption)) {
                            var captionParagraph = host.CreateParagraph();
                            ApplyBlockParagraphFormatting(captionParagraph, quoteDepth, alignment);
                            captionParagraph.AddText(semantic.Caption!);
                        }
                        break;
                    }
                case Omd.TableBlock tb:
                    RenderSharedTableBlockOmd(tb, host, options, document, pageContentWidthPixels);
                    break;
                case Omd.UnorderedListBlock ul: {
                        var list = host.CreateList(WordListStyle.Bulleted);
                        foreach (var item in ul.Items) {
                            int effectiveLevel = listLevel + item.Level;
                            var li = list.AddItem(string.Empty, effectiveLevel);
                            if (item.IsTask) {
                                li.AddCheckBox(item.Checked);
                            }
                            ApplyBlockParagraphFormatting(li, quoteDepth, alignment);
                            ProcessInlinesOmd(item.Content, li, options, document, _currentFootnotes);

                            if (item.AdditionalParagraphs != null && item.AdditionalParagraphs.Count > 0) {
                                foreach (var extra in item.AdditionalParagraphs) {
                                    var li2 = list.AddItem(string.Empty, effectiveLevel);
                                    ApplyBlockParagraphFormatting(li2, quoteDepth, alignment);
                                    ProcessInlinesOmd(extra, li2, options, document, _currentFootnotes);
                                }
                            }

                            if (item.Children != null && item.Children.Count > 0) {
                                foreach (var child in item.Children) {
                                    RenderSharedBlockOmd(child, host, options, document, null, effectiveLevel + 1, quoteDepth, pageContentWidthPixels, alignment);
                                }
                            }
                        }
                        break;
                    }
                case Omd.OrderedListBlock ol: {
                        var list = host.CreateList(WordListStyle.Numbered);
                        if (ol.Start != 1) {
                            list.Numbering.Levels[0].SetStartNumberingValue(ol.Start);
                        }
                        foreach (var item in ol.Items) {
                            int effectiveLevel = listLevel + item.Level;
                            var li = list.AddItem(string.Empty, effectiveLevel);
                            if (item.IsTask) {
                                li.AddCheckBox(item.Checked);
                            }
                            ApplyBlockParagraphFormatting(li, quoteDepth, alignment);
                            ProcessInlinesOmd(item.Content, li, options, document, _currentFootnotes);

                            if (item.AdditionalParagraphs != null && item.AdditionalParagraphs.Count > 0) {
                                foreach (var extra in item.AdditionalParagraphs) {
                                    var li2 = list.AddItem(string.Empty, effectiveLevel);
                                    ApplyBlockParagraphFormatting(li2, quoteDepth, alignment);
                                    ProcessInlinesOmd(extra, li2, options, document, _currentFootnotes);
                                }
                            }

                            if (item.Children != null && item.Children.Count > 0) {
                                foreach (var child in item.Children) {
                                    RenderSharedBlockOmd(child, host, options, document, null, effectiveLevel + 1, quoteDepth, pageContentWidthPixels, alignment);
                                }
                            }
                        }
                        break;
                    }
                case Omd.TocBlock:
                    break;
                case Omd.HtmlCommentBlock comment:
                    if (host.SupportsHtmlInsertion) {
                        host.InsertHtml(comment.Comment);
                    } else {
                        var commentParagraph = host.CreateParagraph();
                        ApplyBlockParagraphFormatting(commentParagraph, quoteDepth, alignment);
                        commentParagraph.AddText(((Omd.IMarkdownBlock)comment).RenderMarkdown());
                    }
                    break;
                case Omd.HtmlRawBlock html:
                    if (host.SupportsHtmlInsertion) {
                        host.InsertHtml(html.Html);
                    } else {
                        var htmlParagraph = host.CreateParagraph();
                        ApplyBlockParagraphFormatting(htmlParagraph, quoteDepth, alignment);
                        htmlParagraph.AddText(((Omd.IMarkdownBlock)html).RenderMarkdown());
                    }
                    break;
                case Omd.HorizontalRuleBlock:
                    if (host.SupportsHorizontalRule) {
                        host.InsertHorizontalRule();
                    } else {
                        var ruleParagraph = host.CreateParagraph();
                        ApplyBlockParagraphFormatting(ruleParagraph, quoteDepth, alignment);
                        ruleParagraph.AddText("---");
                    }
                    break;
                case Omd.DefinitionListBlock dl:
                    foreach (var entry in dl.Entries) {
                        if (entry == null) {
                            continue;
                        }

                        if (string.IsNullOrWhiteSpace(entry.TermMarkdown) && entry.DefinitionBlocks.Count == 0) {
                            continue;
                        }

                        RenderSharedDefinitionListEntryOmd(entry, host, options, document, quoteDepth, pageContentWidthPixels, alignment);
                    }
                    break;
                case Omd.QuoteBlock qb:
                    RenderSharedBlocksOmd(qb.Children, host, options, document, listLevel, quoteDepth + 1, pageContentWidthPixels, alignment);
                    break;
                case Omd.CalloutBlock callout:
                    RenderSharedCalloutBlockOmd(callout, host, options, document, quoteDepth, pageContentWidthPixels, alignment);
                    break;
                case Omd.FootnoteDefinitionBlock:
                    break;
                default: {
                        var fallback = host.CreateParagraph();
                        ApplyBlockParagraphFormatting(fallback, quoteDepth, alignment);
                        fallback.AddText(block.RenderMarkdown());
                        break;
                    }
            }
        }

        private static void ProcessBlockOmd(
            Omd.IMarkdownBlock block,
            WordDocument document,
            MarkdownToWordOptions options,
            WordList? currentList = null,
            int listLevel = 0,
            int quoteDepth = 0,
            double pageContentWidthPixels = 0) {
            RenderSharedBlockOmd(
                block,
                new DocumentWordBlockRenderHost(document),
                options,
                document,
                currentList,
                listLevel,
                quoteDepth,
                pageContentWidthPixels,
                Omd.ColumnAlignment.None);
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
