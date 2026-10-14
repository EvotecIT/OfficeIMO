using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private sealed class DocumentSelectorPresence {
        private readonly HashSet<string> _tags = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _ids = new(StringComparer.Ordinal);
        private readonly HashSet<string> _classes = new(StringComparer.Ordinal);
        private readonly bool _quirksMode;

        internal DocumentSelectorPresence(IHtmlDocument document) {
            _quirksMode = UsesQuirksIdClassMatching(document);
            foreach (IElement element in document.QuerySelectorAll("*")) {
                _tags.Add(element.LocalName ?? element.TagName ?? string.Empty);
                string id = element.Id ?? string.Empty;
                if (id.Length > 0) _ids.Add(SelectorIdentityKey(id, _quirksMode));
                foreach (string name in element.ClassList) _classes.Add(SelectorIdentityKey(name, _quirksMode));
            }
        }

        internal bool CanMatch(string selector) {
            SelectorCandidateKey key = GetSelectorCandidateKey(selector);
            return key.Kind switch {
                SelectorCandidateKind.Tag => _tags.Contains(key.Value),
                SelectorCandidateKind.Id => _ids.Contains(SelectorIdentityKey(key.Value, _quirksMode)),
                SelectorCandidateKind.Class => _classes.Contains(SelectorIdentityKey(key.Value, _quirksMode)),
                _ => true
            };
        }
    }
}
