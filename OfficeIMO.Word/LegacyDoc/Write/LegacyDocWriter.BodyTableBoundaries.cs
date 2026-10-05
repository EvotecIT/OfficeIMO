using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static OpenXmlElement? FindFollowingBodyBlock(OpenXmlElement marker) {
            OpenXmlElement? element = NextBodyElement(marker);
            while (element != null) {
                if (element is SdtBlock control && control.SdtContentBlock != null) {
                    element = control.SdtContentBlock.FirstChild ?? NextBodyElement(control);
                } else if (element is BookmarkStart || element is BookmarkEnd) {
                    element = NextBodyElement(element);
                } else {
                    return element;
                }
            }
            return null;
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

        private static OpenXmlElement? FindPrecedingBodyBlock(OpenXmlElement marker) {
            OpenXmlElement? element = PreviousBodyElement(marker);
            while (element != null) {
                if (element is SdtBlock control && control.SdtContentBlock != null) {
                    element = control.SdtContentBlock.LastChild ?? PreviousBodyElement(control);
                } else if (element is BookmarkStart || element is BookmarkEnd) {
                    element = PreviousBodyElement(element);
                } else {
                    return element;
                }
            }
            return null;
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
