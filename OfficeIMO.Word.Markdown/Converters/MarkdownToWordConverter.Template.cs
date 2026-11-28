using OfficeIMO.OpenXml.Internal;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        public WordDocument ConvertIntoTemplate(Omd.MarkdownDocument markdown, WordDocument document, MarkdownWordExportOptions options, CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(markdown);

            ArgumentNullException.ThrowIfNull(document);

            cancellationToken.ThrowIfCancellationRequested();
            options ??= new MarkdownWordExportOptions();
            options.ResolveDocumentStyles(document);
            options.ApplyDefaults(document);

            var insertion = ResolveTemplateInsertion(document, options);
            var host = new BodyInsertionPointWordBlockRenderHost(document, insertion.Anchor);
            var pageContentWidthPixels = EstimatePageContentWidthPixels(document);
            RenderMarkdownDocument(markdown, host, document, options, pageContentWidthPixels, cancellationToken);

            if (insertion.RemoveAfterRender && insertion.Anchor.Parent.Exists) {
                insertion.Anchor.Remove();
            }

            return document;
        }

        private void RenderMarkdownDocument(
            Omd.MarkdownDocument markdown,
            IWordBlockRenderHost host,
            WordDocument document,
            MarkdownWordExportOptions options,
            double pageContentWidthPixels,
            CancellationToken cancellationToken) {
            var blocks = GetRenderableBlocks(markdown, options);

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
                    alignment: Omd.MarkdownColumnAlignment.None);
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
                    alignment: Omd.MarkdownColumnAlignment.None);
            }
        }

        private static TemplateInsertion ResolveTemplateInsertion(WordDocument document, MarkdownWordExportOptions options) {
            if (!string.IsNullOrWhiteSpace(options.ContentControlTag) || !string.IsNullOrWhiteSpace(options.ContentControlAlias)) {
                var contentControl = FindBlockContentControl(document, options);
                if (contentControl.IsNone) {
                    throw new InvalidOperationException("The requested block content control insertion point was not found.");
                }

                return new TemplateInsertion(contentControl, options.ReplacePlaceholder);
            }

            if (!string.IsNullOrWhiteSpace(options.BookmarkName)) {
                var bookmarkParagraph = FindBookmarkParagraph(document, options.BookmarkName!);
                if (bookmarkParagraph.IsNone) {
                    throw new InvalidOperationException($"Bookmark '{options.BookmarkName}' was not found.");
                }

                return new TemplateInsertion(bookmarkParagraph, options.ReplacePlaceholder);
            }

            OfficeXmlNode body = BodyRoot(document);
            OfficeXmlNode marker = body.CreateElement(WordXml.W, "p");
            OfficeXmlNode sectionProperties = body.WElements("sectPr").LastOrDefault();
            marker = body.InsertBefore(marker, sectionProperties);

            return new TemplateInsertion(marker, removeAfterRender: true);
        }

        /// <summary>The main document body; throws when the document has none, as the model's body accessor does.</summary>
        private static OfficeXmlNode BodyRoot(WordDocument document) {
            OfficeXmlNode body = document.BodyNode();
            return body.Exists ? body : throw new InvalidOperationException("Document body is missing.");
        }

        private static OfficeXmlNode FindBlockContentControl(WordDocument document, MarkdownWordExportOptions options) {
            return BodyRoot(document)
                .ModelDescendants("sdt")
                .Where(sdt => sdt.ContentControlLevel() == WordContentControlLevel.Block)
                .FirstOrDefault(sdt =>
                    MatchesContentControlTag(sdt, options.ContentControlTag) &&
                    MatchesContentControlAlias(sdt, options.ContentControlAlias));
        }

        private static bool MatchesContentControlTag(OfficeXmlNode sdt, string? tag) {
            if (string.IsNullOrWhiteSpace(tag)) {
                return true;
            }

            var actual = sdt.WChild("sdtPr").ChildVal("tag");
            return string.Equals(actual, tag, StringComparison.Ordinal);
        }

        private static bool MatchesContentControlAlias(OfficeXmlNode sdt, string? alias) {
            if (string.IsNullOrWhiteSpace(alias)) {
                return true;
            }

            var actual = sdt.WChild("sdtPr").ChildVal("alias");
            return string.Equals(actual, alias, StringComparison.Ordinal);
        }

        private static OfficeXmlNode FindBookmarkParagraph(WordDocument document, string bookmarkName) {
            var bookmark = BodyRoot(document)
                .ModelDescendants("bookmarkStart")
                .FirstOrDefault(start => string.Equals(start.WAttr("name"), bookmarkName, StringComparison.Ordinal));
            return bookmark.Ancestor(WordXml.W, "p");
        }

        private readonly struct TemplateInsertion {
            public TemplateInsertion(OfficeXmlNode anchor, bool removeAfterRender) {
                Anchor = anchor;
                RemoveAfterRender = removeAfterRender;
            }

            public OfficeXmlNode Anchor { get; }
            public bool RemoveAfterRender { get; }
        }
    }
}
