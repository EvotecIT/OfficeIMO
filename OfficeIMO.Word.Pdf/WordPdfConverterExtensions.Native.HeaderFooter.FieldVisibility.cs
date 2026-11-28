using System.Text;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private const int MaxNativeHeaderFooterFieldScanElements = 65_536;
        private const int MaxNativeHeaderFooterFieldDepth = 256;

        private static bool TryGetVisibleNativeHeaderFooterFieldResult(OfficeXmlNode separator, WordParagraph paragraph,
            ref int scannedElements, out OfficeXmlNode resultRun) {
            resultRun = default;
            OfficeXmlNode story = GetNativeHeaderFooterFieldStory(separator);
            int depth = 1;
            bool hasCachedText = false;
            OfficeXmlNode endingRun = default;
            for (OfficeXmlNode element = NextNativeHeaderFooterFieldElement(separator, story); element.Exists;
                 element = NextNativeHeaderFooterFieldElement(element, story)) {
                if (++scannedElements > MaxNativeHeaderFooterFieldScanElements)
                    throw new System.IO.InvalidDataException("Header or footer field results exceed the supported scan limit.");
                // Field markup the package holds untyped, inside unknown markup, isn't part of the field.
                if (!element.IsModelElement()) continue;
                if (element.IsW("fldChar")) {
                    string? type = element.WAttr("fldCharType");
                    if (type == "begin" && ++depth > MaxNativeHeaderFooterFieldDepth)
                        throw new System.IO.InvalidDataException("Header or footer fields exceed the supported nesting limit.");
                    else if (type == "end" && --depth == 0) {
                        endingRun = element.Ancestor(WordXml.W, "r");
                        break;
                    }
                } else if (depth == 1 && element.IsW("t") && element.Ancestor(WordXml.W, "r") is { Exists: true } run) {
                    hasCachedText = true;
                    if (!IsNativeHiddenHeaderFooterSourceRun(run, paragraph)) {
                        resultRun = run;
                        return true;
                    }
                }
            }
            OfficeXmlNode uncachedRun = endingRun.Exists ? endingRun : separator.Ancestor(WordXml.W, "r");
            if (uncachedRun.IsNone) throw new InvalidOperationException("Sequence contains no elements");
            if (hasCachedText || IsNativeHiddenHeaderFooterSourceRun(uncachedRun, paragraph)) return false;
            resultRun = uncachedRun;
            return true;
        }

        private static bool IsNativeHiddenHeaderFooterSourceRun(OfficeXmlNode run, WordParagraph fallback) {
            OfficeXmlNode sourceParagraph = run.Ancestor(WordXml.W, "p");
            WordParagraph paragraph = sourceParagraph.IsNone || sourceParagraph == fallback.Node()
                ? fallback : fallback._document.ParagraphView(sourceParagraph);
            return IsNativeHiddenHeaderFooterRun(run, paragraph);
        }

        private static string ReadNativeHeaderFooterFieldPrefixCode(OfficeXmlNode begin, OfficeXmlNode paragraph) {
            if (begin.IsNone) return string.Empty;
            OfficeXmlNode story = GetNativeHeaderFooterFieldStory(begin);
            var code = new StringBuilder();
            int depth = 1;
            int scannedElements = 0;
            for (OfficeXmlNode element = NextNativeHeaderFooterFieldElement(begin, story); element.Exists && element != paragraph;
                 element = NextNativeHeaderFooterFieldElement(element, story)) {
                if (++scannedElements > MaxNativeHeaderFooterFieldScanElements)
                    throw new System.IO.InvalidDataException("Header or footer field instructions exceed the supported scan limit.");
                // Field markup the package holds untyped, inside unknown markup, isn't part of the field.
                if (!element.IsModelElement()) continue;
                if (element.IsW("fldChar")) {
                    string? type = element.WAttr("fldCharType");
                    if (type == "begin" && ++depth > MaxNativeHeaderFooterFieldDepth)
                        throw new System.IO.InvalidDataException("Header or footer fields exceed the supported nesting limit.");
                    else if (type == "end" && --depth == 0) break;
                    else if (depth == 1 && type == "separate") break;
                } else if (depth == 1 && element.IsW("instrText")) code.Append(element.InnerText);
            }
            return code.ToString();
        }

        private static OfficeXmlNode GetNativeHeaderFooterFieldStory(OfficeXmlNode element) {
            OfficeXmlNode last = default;
            foreach (OfficeXmlNode ancestor in element.Ancestors()) {
                if (ancestor.IsW("hdr") || ancestor.IsW("ftr") || ancestor.IsW("txbxContent")) return ancestor;
                last = ancestor;
            }
            return last.Exists ? last : element;
        }

        private static OfficeXmlNode NextNativeHeaderFooterFieldElement(OfficeXmlNode element, OfficeXmlNode story) {
            // Stay within this story and mirror the visible-run reader's revision exclusions.
            if (!(element.IsW("txbxContent") || element.IsW("del") || element.IsW("moveFrom")) && element.FirstChild.Exists)
                return element.FirstChild;
            for (OfficeXmlNode current = element; current.Exists && current != story; current = current.Parent) {
                if (current.NextSibling is { Exists: true } next) return next;
            }
            return default;
        }
    }
}
