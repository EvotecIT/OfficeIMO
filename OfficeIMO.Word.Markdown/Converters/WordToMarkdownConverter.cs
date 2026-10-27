using System.Collections.Generic;
using System.Linq;
using System.Threading;
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
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new WordMarkdownExportOptions();

            var markdown = MarkdownDocument.Create();
            Dictionary<string, string>? previousAnchors = _headingAnchorSlugs;
            _headingAnchorSlugs = BuildHeadingAnchorSlugs(document);
            try {
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
            var body = document._wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) return map;
            foreach (var paragraph in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()) {
                string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId == null || !styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) continue;
                string text = string.Concat(paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(item => item.Text)).Trim();
                if (text.Length == 0) continue;
                string slug = MarkdownSlug.Generate(text, MarkdownHeadingIdentifierStyle.OfficeIMO, registry);
                foreach (var bookmark in paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.BookmarkStart>()) {
                    string? name = bookmark.Name?.Value;
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
