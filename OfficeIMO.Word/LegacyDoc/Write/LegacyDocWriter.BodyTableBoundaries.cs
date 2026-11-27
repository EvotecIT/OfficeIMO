namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private sealed class BodyBlockBoundaryCache {
            private readonly Dictionary<LegacyDocElement, LegacyDocElement> _following = new();
            private readonly Dictionary<LegacyDocElement, LegacyDocElement> _preceding = new();

            internal LegacyDocElement FindFollowing(LegacyDocElement marker) => Find(marker, _following, forward: true);
            internal LegacyDocElement FindPreceding(LegacyDocElement marker) => Find(marker, _preceding, forward: false);

            private static LegacyDocElement Find(LegacyDocElement marker,
                Dictionary<LegacyDocElement, LegacyDocElement> cache, bool forward) {
                if (cache.TryGetValue(marker, out LegacyDocElement result)) return result;

                var visited = new List<LegacyDocElement> { marker };
                LegacyDocElement element = forward ? NextBodyElement(marker) : PreviousBodyElement(marker);
                while (element.Exists) {
                    if (cache.TryGetValue(element, out result)) break;
                    LegacyDocElement content = element.Is("sdt") ? element.Child("sdtContent") : LegacyDocElement.None;
                    if (content.Exists) {
                        visited.Add(element);
                        LegacyDocElement inner = forward ? content.FirstElement : LastChildElement(content);
                        element = inner.Exists
                            ? inner
                            : forward ? NextBodyElement(element) : PreviousBodyElement(element);
                    } else if (element.Is("bookmarkStart") || element.Is("bookmarkEnd")) {
                        visited.Add(element);
                        element = forward ? NextBodyElement(element) : PreviousBodyElement(element);
                    } else {
                        result = element;
                        break;
                    }
                }
                foreach (LegacyDocElement visitedElement in visited) cache[visitedElement] = result;
                return result;
            }
        }

        private static LegacyDocElement LastChildElement(LegacyDocElement parent) {
            LegacyDocElement last = LegacyDocElement.None;
            foreach (LegacyDocElement child in parent.Elements()) last = child;
            return last;
        }

        private static LegacyDocElement NextBodyElement(LegacyDocElement element) {
            while (true) {
                LegacyDocElement sibling = element.NextElement;
                if (sibling.Exists) return sibling;
                LegacyDocElement content = element.Parent;
                if (content.Is("sdtContent") && content.Parent.Is("sdt")) {
                    element = content.Parent;
                } else {
                    return LegacyDocElement.None;
                }
            }
        }

        private static LegacyDocElement PreviousBodyElement(LegacyDocElement element) {
            while (true) {
                LegacyDocElement sibling = element.PreviousElement;
                if (sibling.Exists) return sibling;
                LegacyDocElement content = element.Parent;
                if (content.Is("sdtContent") && content.Parent.Is("sdt")) {
                    element = content.Parent;
                } else {
                    return LegacyDocElement.None;
                }
            }
        }
    }
}
