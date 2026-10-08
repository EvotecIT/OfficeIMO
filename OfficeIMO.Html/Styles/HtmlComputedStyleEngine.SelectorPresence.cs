using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private sealed class DocumentSelectorPresence {
        private readonly HashSet<string> _tags = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _ids = new(StringComparer.Ordinal);
        private readonly HashSet<string> _classes = new(StringComparer.Ordinal);

        internal DocumentSelectorPresence(IHtmlDocument document) {
            foreach (IElement element in document.QuerySelectorAll("*")) {
                _tags.Add(element.LocalName ?? element.TagName ?? string.Empty);
                string id = element.Id ?? string.Empty;
                if (id.Length > 0) _ids.Add(id);
                foreach (string name in element.ClassList) _classes.Add(name);
            }
        }

        internal bool CanMatch(string selector) {
            SelectorCandidateKey key = GetSelectorCandidateKey(selector);
            return key.Kind switch {
                SelectorCandidateKind.Tag => _tags.Contains(key.Value),
                SelectorCandidateKind.Id => _ids.Contains(key.Value),
                SelectorCandidateKind.Class => _classes.Contains(key.Value),
                _ => true
            };
        }
    }
}
