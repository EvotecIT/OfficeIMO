using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading;
using System.Threading.Tasks;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        public WordDocument ConvertIntoTemplate(string markdown, WordDocument document, MarkdownToWordTemplateOptions options) {
            return ConvertIntoTemplateAsync(markdown, document, options).GetAwaiter().GetResult();
        }

        public Task<WordDocument> ConvertIntoTemplateAsync(
            string markdown,
            WordDocument document,
            MarkdownToWordTemplateOptions options,
            CancellationToken cancellationToken = default) {
            if (markdown == null) {
                throw new ArgumentNullException(nameof(markdown));
            }

            options ??= new MarkdownToWordTemplateOptions();
            var readerOptions = CreateEffectiveReaderOptions(options);
            var markdownDocument = Omd.MarkdownReader.Parse(markdown, readerOptions);
            return ConvertIntoTemplateAsync(markdownDocument, document, options, cancellationToken);
        }

        public WordDocument ConvertIntoTemplate(Omd.MarkdownDoc markdown, WordDocument document, MarkdownToWordTemplateOptions options) {
            return ConvertIntoTemplateAsync(markdown, document, options).GetAwaiter().GetResult();
        }

        public Task<WordDocument> ConvertIntoTemplateAsync(
            Omd.MarkdownDoc markdown,
            WordDocument document,
            MarkdownToWordTemplateOptions options,
            CancellationToken cancellationToken = default) {
            if (markdown == null) {
                throw new ArgumentNullException(nameof(markdown));
            }

            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            options ??= new MarkdownToWordTemplateOptions();
            options.ApplyDefaults(document);

            var insertion = ResolveTemplateInsertion(document, options);
            var host = new BodyInsertionPointWordBlockRenderHost(document, insertion.Anchor);
            var pageContentWidthPixels = EstimatePageContentWidthPixels(document);
            RenderMarkdownDocument(markdown, host, document, options, pageContentWidthPixels, cancellationToken);

            if (insertion.RemoveAfterRender && insertion.Anchor.Parent != null) {
                insertion.Anchor.Remove();
            }

            return Task.FromResult(document);
        }

        private void RenderMarkdownDocument(
            Omd.MarkdownDoc markdown,
            IWordBlockRenderHost host,
            WordDocument document,
            MarkdownToWordOptions options,
            double pageContentWidthPixels,
            CancellationToken cancellationToken) {
            var blocks = GetRenderableBlocks(markdown);

            _currentFootnotes = blocks
                .OfType<Omd.FootnoteDefinitionBlock>()
                .GroupBy(f => f.Label)
                .ToDictionary(g => g.Key, g => g.Last().Text);

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
        }

        private static TemplateInsertion ResolveTemplateInsertion(WordDocument document, MarkdownToWordTemplateOptions options) {
            if (!string.IsNullOrWhiteSpace(options.ContentControlTag) || !string.IsNullOrWhiteSpace(options.ContentControlAlias)) {
                var contentControl = FindBlockContentControl(document, options);
                if (contentControl == null) {
                    throw new InvalidOperationException("The requested block content control insertion point was not found.");
                }

                return new TemplateInsertion(contentControl, options.ReplacePlaceholder);
            }

            if (!string.IsNullOrWhiteSpace(options.BookmarkName)) {
                var bookmarkParagraph = FindBookmarkParagraph(document, options.BookmarkName!);
                if (bookmarkParagraph == null) {
                    throw new InvalidOperationException($"Bookmark '{options.BookmarkName}' was not found.");
                }

                return new TemplateInsertion(bookmarkParagraph, options.ReplacePlaceholder);
            }

            var marker = new Paragraph();
            var sectionProperties = document.BodyRoot.Elements<SectionProperties>().LastOrDefault();
            if (sectionProperties != null) {
                document.BodyRoot.InsertBefore(marker, sectionProperties);
            } else {
                document.BodyRoot.Append(marker);
            }

            return new TemplateInsertion(marker, removeAfterRender: true);
        }

        private static SdtBlock? FindBlockContentControl(WordDocument document, MarkdownToWordTemplateOptions options) {
            return document.BodyRoot
                .Descendants<SdtBlock>()
                .FirstOrDefault(sdt =>
                    MatchesContentControlTag(sdt, options.ContentControlTag) &&
                    MatchesContentControlAlias(sdt, options.ContentControlAlias));
        }

        private static bool MatchesContentControlTag(SdtBlock sdt, string? tag) {
            if (string.IsNullOrWhiteSpace(tag)) {
                return true;
            }

            var actual = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
            return string.Equals(actual, tag, StringComparison.Ordinal);
        }

        private static bool MatchesContentControlAlias(SdtBlock sdt, string? alias) {
            if (string.IsNullOrWhiteSpace(alias)) {
                return true;
            }

            var actual = sdt.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.Value;
            return string.Equals(actual, alias, StringComparison.Ordinal);
        }

        private static Paragraph? FindBookmarkParagraph(WordDocument document, string bookmarkName) {
            var bookmark = document.BodyRoot
                .Descendants<BookmarkStart>()
                .FirstOrDefault(start => string.Equals(start.Name?.Value, bookmarkName, StringComparison.Ordinal));
            return bookmark?.Ancestors<Paragraph>().FirstOrDefault();
        }

        private readonly struct TemplateInsertion {
            public TemplateInsertion(OpenXmlElement anchor, bool removeAfterRender) {
                Anchor = anchor;
                RemoveAfterRender = removeAfterRender;
            }

            public OpenXmlElement Anchor { get; }
            public bool RemoveAfterRender { get; }
        }
    }
}
