using AngleSharp.Dom;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private bool ShouldStartBodyResourceParagraph(INode node) {
            if (node is not IElement element ||
                (!element.TagName.Equals("img", StringComparison.OrdinalIgnoreCase) &&
                 !element.TagName.Equals("svg", StringComparison.OrdinalIgnoreCase))) {
                return false;
            }

            string display = GetEffectiveDisplay(element);
            return display is "block" or "flow-root" or "list-item" or "table" or "flex" or "grid" ||
                   !HasSurroundingInlineBodyContent(node);
        }

        private static bool HasSurroundingInlineBodyContent(INode node) =>
            HasInlineBodyContent(node.PreviousSibling, previous: true) ||
            HasInlineBodyContent(node.NextSibling, previous: false);

        private static bool HasInlineBodyContent(INode? sibling, bool previous) {
            while (sibling != null) {
                if (sibling is IText text) {
                    if (!string.IsNullOrWhiteSpace(text.Text)) {
                        return true;
                    }
                } else if (sibling is IElement element) {
                    if (_blockTags.Contains(element.TagName)) {
                        return false;
                    }
                    if (!element.TagName.Equals("img", StringComparison.OrdinalIgnoreCase) &&
                        !element.TagName.Equals("svg", StringComparison.OrdinalIgnoreCase)) {
                        return true;
                    }
                }

                sibling = previous ? sibling.PreviousSibling : sibling.NextSibling;
            }

            return false;
        }

        private bool ShouldStartContainerChildParagraph(IElement element) {
            if (!_blockTags.Contains(element.TagName)) {
                return false;
            }
            if (!element.TagName.Equals("code", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            string display = GetEffectiveDisplay(element);
            return display is "block" or "flow-root" or "list-item" or "table" or "flex" or "grid";
        }

        private string GetEffectiveDisplay(IElement element) {
            var declarations = CollectCssDeclarations(element, inheritedOnly: false);
            return declarations.TryGetValue("display", out var display)
                ? display.Value.Trim().ToLowerInvariant()
                : string.Empty;
        }
    }
}
