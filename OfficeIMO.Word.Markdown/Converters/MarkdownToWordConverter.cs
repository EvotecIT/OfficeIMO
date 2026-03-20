using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Markdown.Html;
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
                var defaults = new Omd.MarkdownReaderOptions {
                    BaseUri = options.BaseUri,
                    PreferNarrativeSingleLineDefinitions = options.PreferNarrativeSingleLineDefinitions
                };
                WordMarkdownSemanticBlocks.ConfigureReaderOptions(defaults);
                return defaults;
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
            WordMarkdownSemanticBlocks.ConfigureReaderOptions(effective);
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

            if (omd.DocumentHeader != null) {
                ProcessBlockOmd(omd.DocumentHeader, document, options, quoteDepth: 0, pageContentWidthPixels: pageContentWidthPixels);
            }

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
            if (markdown.DocumentHeader != null) {
                RenderSharedBlockOmd(
                    markdown.DocumentHeader,
                    host,
                    options,
                    document,
                    currentList: null,
                    listLevel: 0,
                    quoteDepth: 0,
                    pageContentWidthPixels: pageContentWidthPixels,
                    alignment: Omd.ColumnAlignment.None);
            }

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

        private sealed class HeaderFooterWordBlockRenderHost : IWordBlockRenderHost {
            private readonly WordHeaderFooter _headerFooter;

            public HeaderFooterWordBlockRenderHost(WordHeaderFooter headerFooter) {
                _headerFooter = headerFooter ?? throw new ArgumentNullException(nameof(headerFooter));
            }

            public WordParagraph CreateParagraph() => _headerFooter.AddParagraph(string.Empty);
            public WordList CreateList(WordListStyle style) => _headerFooter.AddList(style);
            public WordTable CreateTable(int rows, int columns) => _headerFooter.AddTable(rows, columns);
            public bool SupportsHtmlInsertion => false;
            public void InsertHtml(string html) { }
            public bool SupportsHorizontalRule => true;
            public void InsertHorizontalRule() => _headerFooter.AddHorizontalLine();
        }

        private static void ApplyBlockParagraphFormatting(WordParagraph paragraph, int quoteDepth, Omd.ColumnAlignment alignment) {
            if (quoteDepth > 0) {
                paragraph.IndentationBefore = IndentTwipsPerLevel * quoteDepth;
            }

            ApplyAlignment(alignment, paragraph);
        }

        private static bool TryRenderHtmlFallbackViaMarkdownAst(
            string html,
            IWordBlockRenderHost host,
            MarkdownToWordOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.ColumnAlignment alignment) {
            if (string.IsNullOrWhiteSpace(html)) {
                return true;
            }

            Omd.MarkdownDoc htmlDocument;
            try {
                var htmlOptions = HtmlToMarkdownOptions.CreateOfficeIMOProfile();
                htmlOptions.PreserveUnsupportedBlocks = false;
                htmlOptions.PreserveUnsupportedInlineHtml = false;
                htmlDocument = html.LoadFromHtml(htmlOptions);
            } catch {
                return false;
            }

            if (htmlDocument.DocumentHeader != null) {
                RenderSharedBlockOmd(
                    htmlDocument.DocumentHeader,
                    host,
                    options,
                    document,
                    currentList: null,
                    listLevel: 0,
                    quoteDepth: quoteDepth,
                    pageContentWidthPixels: pageContentWidthPixels,
                    alignment: alignment);
            }

            var renderedAny = false;
            foreach (var block in htmlDocument.Blocks) {
                if (block == null) {
                    continue;
                }

                renderedAny = true;
                RenderSharedBlockOmd(
                    block,
                    host,
                    options,
                    document,
                    currentList: null,
                    listLevel: 0,
                    quoteDepth: quoteDepth,
                    pageContentWidthPixels: pageContentWidthPixels,
                    alignment: alignment);
            }

            return renderedAny;
        }

        private static bool TryRenderWordHeaderFooterSemanticBlock(
            Omd.SemanticFencedBlock block,
            IWordBlockRenderHost currentHost,
            MarkdownToWordOptions options,
            WordDocument document,
            double pageContentWidthPixels) {
            if (block == null || currentHost is not DocumentWordBlockRenderHost) {
                return false;
            }

            if (!TryResolveWordHeaderFooterTarget(block, document, options, out var target)) {
                return false;
            }

            var targetHost = new HeaderFooterWordBlockRenderHost(target);
            if (!string.IsNullOrWhiteSpace(block.Content)) {
                var readerOptions = CreateEffectiveReaderOptions(options);
                readerOptions.FrontMatter = false;
                var fragment = Omd.MarkdownReader.Parse(block.Content, readerOptions);

                if (fragment.DocumentHeader != null) {
                    RenderSharedBlockOmd(
                        fragment.DocumentHeader,
                        targetHost,
                        options,
                        document,
                        currentList: null,
                        listLevel: 0,
                        quoteDepth: 0,
                        pageContentWidthPixels: pageContentWidthPixels,
                        alignment: Omd.ColumnAlignment.None);
                }

                foreach (var nested in fragment.Blocks ?? Array.Empty<Omd.IMarkdownBlock>()) {
                    RenderSharedBlockOmd(
                        nested,
                        targetHost,
                        options,
                        document,
                        currentList: null,
                        listLevel: 0,
                        quoteDepth: 0,
                        pageContentWidthPixels: pageContentWidthPixels,
                        alignment: Omd.ColumnAlignment.None);
                }
            }

            if (!string.IsNullOrWhiteSpace(block.Caption)) {
                target.AddParagraph(block.Caption!);
            }

            return true;
        }

        private static bool TryResolveWordHeaderFooterTarget(
            Omd.SemanticFencedBlock block,
            WordDocument document,
            MarkdownToWordOptions options,
            out WordHeaderFooter target) {
            target = null!;
            if (block == null) {
                return false;
            }

            bool isHeader;
            if (string.Equals(block.SemanticKind, WordMarkdownSemanticBlocks.HeaderSemanticKind, StringComparison.OrdinalIgnoreCase)) {
                isHeader = true;
            } else if (string.Equals(block.SemanticKind, WordMarkdownSemanticBlocks.FooterSemanticKind, StringComparison.OrdinalIgnoreCase)) {
                isHeader = false;
            } else {
                return false;
            }

            int sectionNumber = 1;
            if (block.FenceInfo.TryGetInt32Attribute("section", out var parsedSection) && parsedSection > 0) {
                sectionNumber = parsedSection;
            }

            if (sectionNumber != 1) {
                options.OnWarning?.Invoke($"Semantic {block.SemanticKind} block requested section {sectionNumber}, but MarkdownToWord currently restores headers and footers only for section 1.");
            }

            var slot = block.FenceInfo.GetAttribute("slot");
            var type = ResolveHeaderFooterType(slot, options, block);
            target = isHeader
                ? document.Sections[0].GetOrCreateHeader(type)
                : document.Sections[0].GetOrCreateFooter(type);
            return true;
        }

        private static HeaderFooterValues ResolveHeaderFooterType(
            string? slot,
            MarkdownToWordOptions options,
            Omd.SemanticFencedBlock block) {
            if (string.IsNullOrWhiteSpace(slot)) {
                return HeaderFooterValues.Default;
            }

            var normalizedSlot = (slot ?? string.Empty).Trim().ToLowerInvariant();
            return normalizedSlot switch {
                "default" => HeaderFooterValues.Default,
                "odd" => HeaderFooterValues.Default,
                "first" => HeaderFooterValues.First,
                "even" => HeaderFooterValues.Even,
                _ => WarnAndReturnDefault(normalizedSlot, options, block)
            };
        }

        private static HeaderFooterValues WarnAndReturnDefault(
            string slot,
            MarkdownToWordOptions options,
            Omd.SemanticFencedBlock block) {
            options.OnWarning?.Invoke($"Semantic {block.SemanticKind} block requested unsupported slot '{slot}'. Falling back to default header or footer.");
            return HeaderFooterValues.Default;
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

        private sealed class BlockRenderer : Omd.MarkdownVisitor {
            private readonly IWordBlockRenderHost _host;
            private readonly MarkdownToWordOptions _options;
            private readonly WordDocument _document;
            private readonly int _listLevel;
            private readonly int _quoteDepth;
            private readonly double _pageContentWidthPixels;
            private readonly Omd.ColumnAlignment _alignment;

            public BlockRenderer(
                IWordBlockRenderHost host,
                MarkdownToWordOptions options,
                WordDocument document,
                int listLevel,
                int quoteDepth,
                double pageContentWidthPixels,
                Omd.ColumnAlignment alignment) {
                _host = host ?? throw new ArgumentNullException(nameof(host));
                _options = options ?? throw new ArgumentNullException(nameof(options));
                _document = document ?? throw new ArgumentNullException(nameof(document));
                _listLevel = listLevel;
                _quoteDepth = quoteDepth;
                _pageContentWidthPixels = pageContentWidthPixels;
                _alignment = alignment;
            }

            public void Render(Omd.IMarkdownBlock block) {
                if (block == null) {
                    return;
                }

                if (block is Omd.MarkdownObject markdownObject) {
                    Visit(markdownObject);
                } else {
                    RenderFallback(block);
                }
            }

            private void RenderNested(
                Omd.IMarkdownBlock block,
                int? listLevel = null,
                int? quoteDepth = null,
                double? pageContentWidthPixels = null,
                Omd.ColumnAlignment? alignment = null) {
                new BlockRenderer(
                    _host,
                    _options,
                    _document,
                    listLevel ?? _listLevel,
                    quoteDepth ?? _quoteDepth,
                    pageContentWidthPixels ?? _pageContentWidthPixels,
                    alignment ?? _alignment)
                    .Render(block);
            }

            private void RenderFallback(Omd.IMarkdownBlock block) {
                var fallback = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(fallback, _quoteDepth, _alignment);
                fallback.AddText(block.RenderMarkdown());
            }

            protected override void VisitBlock(Omd.MarkdownBlock block) {
                if (block is Omd.IMarkdownBlock markdownBlock) {
                    RenderFallback(markdownBlock);
                }
            }

            protected override void VisitHeadingBlock(Omd.HeadingBlock block) {
                var headingParagraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(headingParagraph, _quoteDepth, _alignment);
                ProcessInlinesOmd(block.Inlines, headingParagraph, _options, _document, _currentFootnotes);
                headingParagraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(block.Level);
            }

            protected override void VisitParagraphBlock(Omd.ParagraphBlock block) {
                var paragraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                ProcessInlinesOmd(block.Inlines, paragraph, _options, _document, _currentFootnotes);
            }

            protected override void VisitImageBlock(Omd.ImageBlock block) {
                var paragraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                var pathOrUrl = block.Path ?? string.Empty;
                var contextWidthLimit = ResolveContextWidthLimitPixels(_options.ImageLayout, _pageContentWidthPixels, _listLevel, _quoteDepth);

                if (System.IO.File.Exists(pathOrUrl)) {
                    if (_options.AllowLocalImages && LocalPathAllowed(pathOrUrl, _options)) {
                        double? naturalW = null;
                        double? naturalH = null;
                        if (TryGetImageDimensionsFromFile(pathOrUrl, out var fileW, out var fileH)) {
                            naturalW = fileW;
                            naturalH = fileH;
                        }

                        ResolveImageDimensions(
                            _options,
                            source: pathOrUrl,
                            context: "block-local",
                            requestedWidth: block.Width,
                            requestedHeight: block.Height,
                            naturalWidth: naturalW,
                            naturalHeight: naturalH,
                            pageContentWidthPixels: _pageContentWidthPixels,
                            contextWidthLimitPixels: contextWidthLimit,
                            out var finalW,
                            out var finalH,
                            out _);

                        paragraph.AddImage(pathOrUrl, finalW, finalH, description: block.Alt ?? string.Empty);
                    } else {
                        var text = paragraph.AddText(block.Alt ?? System.IO.Path.GetFileName(pathOrUrl));
                        var defaultFont = ResolveDefaultFontFamily(_options);
                        if (!string.IsNullOrEmpty(defaultFont)) {
                            text.SetFontFamily(defaultFont!);
                        }
                    }
                } else if (System.Uri.TryCreate(pathOrUrl, System.UriKind.Absolute, out var uri)) {
                    if (_options.AllowedImageSchemes.Contains(uri.Scheme) &&
                        (_options.ImageUrlValidator == null || _options.ImageUrlValidator(uri))) {
                        if (_options.AllowRemoteImages) {
                            try {
                                var bytes = DownloadRemoteImageBytes(uri, _options);
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
                                    _options,
                                    source: uri.ToString(),
                                    context: "block-remote",
                                    requestedWidth: block.Width,
                                    requestedHeight: block.Height,
                                    naturalWidth: naturalW,
                                    naturalHeight: naturalH,
                                    pageContentWidthPixels: _pageContentWidthPixels,
                                    contextWidthLimitPixels: contextWidthLimit,
                                    out var finalW,
                                    out var finalH,
                                    out _);

                                using var stream = new System.IO.MemoryStream(bytes, writable: false);
                                paragraph.AddImage(stream, fileName, finalW, finalH, description: block.Alt ?? string.Empty);
                            } catch (Exception ex) {
                                _options.OnWarning?.Invoke($"Remote image '{uri}' could not be downloaded. {ex.Message}");
                                if (_options.FallbackRemoteImagesToHyperlinks) {
                                    paragraph.AddHyperLink(block.Alt ?? uri.ToString(), uri);
                                }
                            }
                        } else if (_options.FallbackRemoteImagesToHyperlinks) {
                            paragraph.AddHyperLink(block.Alt ?? uri.ToString(), uri);
                        }
                    } else if (_options.FallbackRemoteImagesToHyperlinks) {
                        paragraph.AddHyperLink(block.Alt ?? uri.ToString(), uri);
                    }
                } else {
                    var text = paragraph.AddText(block.Alt ?? pathOrUrl);
                    var defaultFont = ResolveDefaultFontFamily(_options);
                    if (!string.IsNullOrEmpty(defaultFont)) {
                        text.SetFontFamily(defaultFont!);
                    }
                }

                if (!string.IsNullOrWhiteSpace(block.Caption)) {
                    var captionParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(captionParagraph, _quoteDepth, _alignment);
                    captionParagraph.AddText(block.Caption!);
                }
            }

            protected override void VisitCodeBlock(Omd.CodeBlock block) {
                var codeParagraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(codeParagraph, _quoteDepth, _alignment);
                var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";
                codeParagraph.AddFormattedText(block.Content ?? string.Empty).SetFontFamily(monoFont);
                if (!string.IsNullOrWhiteSpace(block.Caption)) {
                    var captionParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(captionParagraph, _quoteDepth, _alignment);
                    captionParagraph.AddText(block.Caption!);
                }
            }

            protected override void VisitSemanticFencedBlock(Omd.SemanticFencedBlock block) {
                if (TryRenderWordHeaderFooterSemanticBlock(block, _host, _options, _document, _pageContentWidthPixels)) {
                    return;
                }

                var paragraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";
                paragraph.AddFormattedText(block.Content ?? string.Empty).SetFontFamily(monoFont);
                if (!string.IsNullOrWhiteSpace(block.Caption)) {
                    var captionParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(captionParagraph, _quoteDepth, _alignment);
                    captionParagraph.AddText(block.Caption!);
                }
            }

            protected override void VisitTableBlock(Omd.TableBlock block) =>
                RenderSharedTableBlockOmd(block, _host, _options, _document, _pageContentWidthPixels);

            protected override void VisitUnorderedListBlock(Omd.UnorderedListBlock block) =>
                RenderListBlock(block.Items, WordListStyle.Bulleted, startNumber: null);

            protected override void VisitOrderedListBlock(Omd.OrderedListBlock block) =>
                RenderListBlock(block.Items, WordListStyle.Numbered, block.Start);

            protected override void VisitTocBlock(Omd.TocBlock block) { }

            protected override void VisitHtmlCommentBlock(Omd.HtmlCommentBlock block) {
                if (_host.SupportsHtmlInsertion) {
                    _host.InsertHtml(block.Comment);
                }
            }

            protected override void VisitHtmlRawBlock(Omd.HtmlRawBlock block) {
                if (_host.SupportsHtmlInsertion) {
                    _host.InsertHtml(block.Html);
                } else if (TryRenderHtmlFallbackViaMarkdownAst(block.Html, _host, _options, _document, _quoteDepth, _pageContentWidthPixels, _alignment)) {
                    return;
                } else {
                    var htmlParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(htmlParagraph, _quoteDepth, _alignment);
                    htmlParagraph.AddText(((Omd.IMarkdownBlock)block).RenderMarkdown());
                }
            }

            protected override void VisitHorizontalRuleBlock(Omd.HorizontalRuleBlock block) {
                if (_host.SupportsHorizontalRule) {
                    _host.InsertHorizontalRule();
                } else {
                    var ruleParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(ruleParagraph, _quoteDepth, _alignment);
                    ruleParagraph.AddText("---");
                }
            }

            protected override void VisitDefinitionListBlock(Omd.DefinitionListBlock block) {
                foreach (var entry in block.Entries) {
                    if (entry == null) {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(entry.TermMarkdown) && entry.DefinitionBlocks.Count == 0) {
                        continue;
                    }

                    RenderSharedDefinitionListEntryOmd(entry, _host, _options, _document, _quoteDepth, _pageContentWidthPixels, _alignment);
                }
            }

            protected override void VisitQuoteBlock(Omd.QuoteBlock block) {
                foreach (var child in block.Children) {
                    RenderNested(child, quoteDepth: _quoteDepth + 1);
                }
            }

            protected override void VisitCalloutBlock(Omd.CalloutBlock block) =>
                RenderSharedCalloutBlockOmd(block, _host, _options, _document, _quoteDepth, _pageContentWidthPixels, _alignment);

            protected override void VisitFootnoteDefinitionBlock(Omd.FootnoteDefinitionBlock block) { }

            protected override void VisitDetailsBlock(Omd.DetailsBlock block) {
                if (block.Summary != null) {
                    var summaryParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(summaryParagraph, _quoteDepth, _alignment);
                    ProcessInlinesOmd(block.Summary.Inlines, summaryParagraph, _options, _document, _currentFootnotes);
                    foreach (var run in summaryParagraph.GetRuns()) {
                        run.SetBold();
                    }
                }

                foreach (var child in block.ChildBlocks) {
                    RenderNested(child, quoteDepth: _quoteDepth + 1);
                }
            }

            protected override void VisitSummaryBlock(Omd.SummaryBlock block) {
                var summaryParagraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(summaryParagraph, _quoteDepth, _alignment);
                ProcessInlinesOmd(block.Inlines, summaryParagraph, _options, _document, _currentFootnotes);
                foreach (var run in summaryParagraph.GetRuns()) {
                    run.SetBold();
                }
            }

            protected override void VisitFrontMatterBlock(Omd.FrontMatterBlock block) {
                var lines = block.Render().Replace("\r", string.Empty).Split('\n');
                var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";

                for (int i = 0; i < lines.Length; i++) {
                    var paragraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                    paragraph.AddFormattedText(lines[i]).SetFontFamily(monoFont);
                }
            }

            private void RenderListBlock(IReadOnlyList<Omd.ListItem> items, WordListStyle style, int? startNumber) {
                var list = _host.CreateList(style);
                if (startNumber.HasValue && startNumber.Value != 1) {
                    list.Numbering.Levels[0].SetStartNumberingValue(startNumber.Value);
                }

                foreach (var item in items) {
                    var effectiveLevel = _listLevel + item.Level;
                    var firstParagraph = true;
                    var blockChildren = item.BlockChildren;

                    for (int i = 0; i < blockChildren.Count; i++) {
                        if (blockChildren[i] is Omd.ParagraphBlock paragraph) {
                            var listItemParagraph = list.AddItem((string?)null, effectiveLevel);
                            if (firstParagraph && item.IsTask) {
                                listItemParagraph.AddCheckBox(item.Checked);
                            }

                            ApplyBlockParagraphFormatting(listItemParagraph, _quoteDepth, _alignment);
                            ProcessInlinesOmd(paragraph.Inlines, listItemParagraph, _options, _document, _currentFootnotes);
                            firstParagraph = false;
                            continue;
                        }

                        RenderNested(blockChildren[i], listLevel: effectiveLevel + 1);
                    }
                }
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
            _ = currentList;
            new BlockRenderer(host, options, document, listLevel, quoteDepth, pageContentWidthPixels, alignment).Render(block);
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
