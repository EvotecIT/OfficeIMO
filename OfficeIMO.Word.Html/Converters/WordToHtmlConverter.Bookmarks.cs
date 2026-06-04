using AngleSharp.Dom;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private void ApplyBookmarkId(IElement element, WordParagraph paragraph) {
            if (!TryGetHtmlBookmarkName(paragraph, out var name)) {
                return;
            }

            element.SetAttribute("id", name);
        }

        private bool TryGetHtmlBookmarkName(WordParagraph paragraph, out string name) {
            name = string.Empty;
            if (!paragraph.IsBookmark || paragraph.Bookmark == null) {
                return false;
            }

            name = paragraph.Bookmark.Name ?? string.Empty;
            if (string.IsNullOrWhiteSpace(name)) {
                return false;
            }

            var parts = name.Split(new[] { ':' }, 2);
            if (parts.Length == 2 && IsStructuralTag(parts[0])) {
                return false;
            }

            return true;
        }
    }
}
