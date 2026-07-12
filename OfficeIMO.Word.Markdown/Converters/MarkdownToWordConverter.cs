using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
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


        // Current footnote definitions map; scoped to this per-conversion converter instance.
        private IReadOnlyDictionary<string, string>? _currentFootnotes;

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
            var blocks = GetRenderableBlocks(omd);
            // Build footnote definitions map for this document
            _currentFootnotes = blocks
                .OfType<Omd.FootnoteDefinitionBlock>()
                .GroupBy(f => f.Label)
                .ToDictionary(g => g.Key, g => g.Last().Text);

            if (omd.DocumentHeader != null) {
                ProcessBlockOmd(omd.DocumentHeader, document, options, quoteDepth: 0, pageContentWidthPixels: pageContentWidthPixels);
            }

            foreach (var block in blocks) {
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
            var blocks = GetRenderableBlocks(markdown);

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

        private static void RemoveDuplicateNativeTocTitleHeadings(List<Omd.IMarkdownBlock> blocks) {
            for (int i = 1; i < blocks.Count; i++) {
                if (blocks[i] is not Omd.TocBlock toc ||
                    !toc.TitleHeadingAlreadyRendered ||
                    !toc.IncludeTitle ||
                    string.IsNullOrWhiteSpace(toc.Title) ||
                    toc.Scope != Omd.TocScope.Document ||
                    blocks[i - 1] is not Omd.HeadingBlock heading) {
                    continue;
                }

                int titleLevel = toc.TitleLevel < 1 ? 1 : (toc.TitleLevel > 6 ? 6 : toc.TitleLevel);
                if (heading.Level == titleLevel &&
                    string.Equals(heading.Text.Trim(), toc.Title.Trim(), StringComparison.Ordinal)) {
                    blocks.RemoveAt(i - 1);
                    i--;
                }
            }
        }

        private static IReadOnlyList<Omd.IMarkdownBlock> GetRenderableBlocks(Omd.MarkdownDoc markdown) {
            var blocks = markdown.GetBlocksAndHeadingSlugs().Blocks ?? new List<Omd.IMarkdownBlock>();
            RemoveDuplicateNativeTocTitleHeadings(blocks);
            return blocks;
        }

    }
}
