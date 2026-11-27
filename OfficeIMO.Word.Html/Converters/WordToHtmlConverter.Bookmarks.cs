using AngleSharp.Dom;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private void AppendAdditionalBookmarkAnchors(IElement parent, WordParagraph paragraph, WordHtmlExportOptions options) {
            var emitted = new HashSet<string>(StringComparer.Ordinal);
            for (IElement? container = parent; container != null; container = container.ParentElement) {
                string? primary = container.GetAttribute("id");
                if (primary != null) emitted.Add(primary);
            }
            foreach (OfficeXmlNode bookmark in paragraph.Node().ModelDescendants("bookmarkStart")) {
                string? name = bookmark.WAttr("name");
                if (string.IsNullOrWhiteSpace(name) || !emitted.Add(name!)) continue;
                string[] parts = name!.Split(new[] { ':' }, 2);
                if (parts.Length == 2 && IsStructuralTag(parts[0])) continue;
                IElement anchor = CreateOutputElement(parent.Owner!, "span");
                SetOutputAttribute(anchor, "id", name, "Bookmark:id");
                parent.AppendChild(anchor);
                AddExportDiagnostic(options, "bookmark-position-projected",
                    "An additional or nested Word bookmark targets its containing HTML paragraph; its text remains in the static projection.",
                    OfficeConversionLossKind.Approximation);
            }
        }

        private void ApplyBookmarkId(IElement element, WordParagraph paragraph) {
            if (!TryGetHtmlBookmarkName(paragraph, out var name)) {
                return;
            }

            SetOutputAttribute(element, "id", name, "Bookmark:id");
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
