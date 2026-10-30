using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private sealed class BodyBlockBoundaryCache {
            private readonly Dictionary<OpenXmlElement, OpenXmlElement?> _following = new();
            private readonly Dictionary<OpenXmlElement, OpenXmlElement?> _preceding = new();

            internal OpenXmlElement? FindFollowing(OpenXmlElement marker) => Find(marker, _following, forward: true);
            internal OpenXmlElement? FindPreceding(OpenXmlElement marker) => Find(marker, _preceding, forward: false);

            private static OpenXmlElement? Find(OpenXmlElement marker,
                Dictionary<OpenXmlElement, OpenXmlElement?> cache, bool forward) {
                if (cache.TryGetValue(marker, out OpenXmlElement? result)) return result;

                var visited = new List<OpenXmlElement> { marker };
                OpenXmlElement? element = forward ? NextBodyElement(marker) : PreviousBodyElement(marker);
                while (element != null) {
                    if (cache.TryGetValue(element, out result)) break;
                    if (element is SdtBlock control && control.SdtContentBlock != null) {
                        visited.Add(element);
                        element = forward
                            ? control.SdtContentBlock.FirstChild ?? NextBodyElement(control)
                            : control.SdtContentBlock.LastChild ?? PreviousBodyElement(control);
                    } else if (element is BookmarkStart || element is BookmarkEnd) {
                        visited.Add(element);
                        element = forward ? NextBodyElement(element) : PreviousBodyElement(element);
                    } else {
                        result = element;
                        break;
                    }
                }
                foreach (OpenXmlElement visitedElement in visited) cache[visitedElement] = result;
                return result;
            }
        }

        private static OpenXmlElement? NextBodyElement(OpenXmlElement element) {
            while (true) {
                OpenXmlElement? sibling = element.NextSibling();
                if (sibling != null) return sibling;
                if (element.Parent is SdtContentBlock content && content.Parent is SdtBlock control) {
                    element = control;
                } else {
                    return null;
                }
            }
        }

        private static OpenXmlElement? PreviousBodyElement(OpenXmlElement element) {
            while (true) {
                OpenXmlElement? sibling = element.PreviousSibling();
                if (sibling != null) return sibling;
                if (element.Parent is SdtContentBlock content && content.Parent is SdtBlock control) {
                    element = control;
                } else {
                    return null;
                }
            }
        }
    }
}
