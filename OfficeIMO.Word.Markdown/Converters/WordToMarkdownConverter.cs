using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// IMPLEMENTATION GUIDELINES:
    /// 1. Read document content using OfficeIMO.Word API:
    ///    - document.Paragraphs for text content
    ///    - paragraph.Style to determine heading levels
    ///    - document.Lists for bullet/numbered lists
    ///    - document.Tables for tables
    /// 2. Convert OfficeIMO.Word elements to Markdown syntax:
    ///    - WordParagraphStyles.Heading1 -> # Heading
    ///    - WordParagraphStyles.Heading2 -> ## Heading
    ///    - Bold text -> **text**
    ///    - Italic text -> *text*
    ///    - Lists -> - item or 1. item
    ///    - Tables -> | col1 | col2 |
    /// 3. Check paragraph.IsListItem to identify list items
    /// 4. Use paragraph.Bold, paragraph.Italic for inline formatting
    /// </summary>
    internal partial class WordToMarkdownConverter {
        public MarkdownDocument ConvertToDocument(WordDocument document, WordMarkdownExportOptions options, CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(document);
            options ??= new WordMarkdownExportOptions();

            var markdown = MarkdownDocument.Create();
            Dictionary<string, string>? previousAnchors = _headingAnchorSlugs;
            _headingAnchorSlugs = BuildHeadingAnchorSlugs(document);
            try {
                // Table cells, running content and notes resolve list info per paragraph, and
                // each appended block would otherwise rebind the whole Markdown object tree.
                using (WordDocumentTraversal.UseListNumberingCatalog(document))
                using (markdown.DeferObjectTreeBinding())
                using (WordComplexFieldRunVisibility.BeginConversionScope())
                    BuildMarkdownDocument(document, markdown, options, cancellationToken);
            } finally {
                _headingAnchorSlugs = previousAnchors;
            }
            return markdown;
        }

        // Bookmarks on headings become the Markdown heading slugs, so in-document links keep resolving
        // after export (a Word bookmark "quick_start" on "Quick start" becomes "#quick-start").
        [ThreadStatic] private static Dictionary<string, string>? _headingAnchorSlugs;

        private static Dictionary<string, string> BuildHeadingAnchorSlugs(WordDocument document) {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            var registry = MarkdownSlug.CreateRegistry();
            OfficeXmlNode body = document.BodyNode();
            if (body.IsNone) return map;
            foreach (OfficeXmlNode paragraph in body.DescendantElements(WordXml.W, "p")) {
                string? styleId = paragraph.WChild("pPr").ChildVal("pStyle");
                if (document.GetHeadingLevel(styleId) == 0) continue;
                // Headings inside table cells are not Markdown headings, so they take no slug.
                if (paragraph.Ancestor(WordXml.W, "tc").Exists) continue;
                string text = string.Concat(paragraph.DescendantElements(WordXml.W, "t").Select(item => item.InnerText)).Trim();
                if (text.Length == 0) continue;
                string slug = MarkdownSlug.Generate(text, MarkdownHeadingIdentifierStyle.OfficeIMO, registry);
                foreach (OfficeXmlNode bookmark in paragraph.DescendantElements(WordXml.W, "bookmarkStart")) {
                    string? name = bookmark.WAttr("name");
                    if (!string.IsNullOrEmpty(name) && !map.ContainsKey(name!)) map[name!] = slug;
                }
            }
            return map;
        }

        /// <summary>Returns the Markdown link target for a Word hyperlink, including in-document bookmark links.</summary>
        internal static string? GetHyperlinkTarget(WordHyperLink? hyperlink) {
            if (hyperlink == null) return null;
            if (hyperlink.Uri != null) return BuildHyperlinkUrl(hyperlink.Uri);
            string? anchor = hyperlink.Anchor;
            if (string.IsNullOrWhiteSpace(anchor)) return null;
            return "#" + (_headingAnchorSlugs != null && _headingAnchorSlugs.TryGetValue(anchor!, out var slug) ? slug : anchor);
        }

        private static string NormalizeMarkdownLineEndings(string markdown) {
            return string.IsNullOrEmpty(markdown)
                ? string.Empty
                : markdown.Replace("\r\n", "\n").TrimEnd('\n');
        }
    }
}
