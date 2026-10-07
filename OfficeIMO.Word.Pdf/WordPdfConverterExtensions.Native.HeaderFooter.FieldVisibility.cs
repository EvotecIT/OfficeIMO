using System.Text;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool HasVisibleNativeHeaderFooterFieldResult(W.FieldChar separator, WordParagraph paragraph) {
            OpenXmlElement story = GetNativeHeaderFooterFieldStory(separator);
            int depth = 1;
            bool hasCachedText = false;
            W.Run? endingRun = null;
            for (OpenXmlElement? element = NextNativeHeaderFooterFieldElement(separator, story); element != null;
                 element = NextNativeHeaderFooterFieldElement(element, story)) {
                if (element is W.FieldChar marker) {
                    if (marker.FieldCharType?.Value == W.FieldCharValues.Begin) depth++;
                    else if (marker.FieldCharType?.Value == W.FieldCharValues.End && --depth == 0) {
                        endingRun = marker.Ancestors<W.Run>().FirstOrDefault();
                        break;
                    }
                } else if (depth == 1 && element is W.Text && element.Ancestors<W.Run>().FirstOrDefault() is W.Run run) {
                    hasCachedText = true;
                    if (!IsNativeHiddenHeaderFooterSourceRun(run, paragraph)) return true;
                }
            }
            return !hasCachedText && !IsNativeHiddenHeaderFooterSourceRun(
                endingRun ?? separator.Ancestors<W.Run>().First(), paragraph);
        }

        private static bool IsNativeHiddenHeaderFooterSourceRun(W.Run run, WordParagraph fallback) {
            W.Paragraph? sourceParagraph = run.Ancestors<W.Paragraph>().FirstOrDefault();
            WordParagraph paragraph = sourceParagraph == null || ReferenceEquals(sourceParagraph, fallback._paragraph)
                ? fallback : new WordParagraph(fallback._document, sourceParagraph);
            return IsNativeHiddenHeaderFooterRun(run, paragraph);
        }

        private static string ReadNativeHeaderFooterFieldPrefixCode(W.FieldChar? begin, W.Paragraph paragraph) {
            if (begin == null) return string.Empty;
            OpenXmlElement story = GetNativeHeaderFooterFieldStory(begin);
            var code = new StringBuilder();
            int depth = 1;
            for (OpenXmlElement? element = NextNativeHeaderFooterFieldElement(begin, story); element != null && !ReferenceEquals(element, paragraph);
                 element = NextNativeHeaderFooterFieldElement(element, story)) {
                if (element is W.FieldChar marker) {
                    if (marker.FieldCharType?.Value == W.FieldCharValues.Begin) depth++;
                    else if (marker.FieldCharType?.Value == W.FieldCharValues.End && --depth == 0) break;
                    else if (depth == 1 && marker.FieldCharType?.Value == W.FieldCharValues.Separate) break;
                } else if (depth == 1 && element is W.FieldCode instruction) code.Append(instruction.Text);
            }
            return code.ToString();
        }

        private static OpenXmlElement GetNativeHeaderFooterFieldStory(OpenXmlElement element) =>
            element.Ancestors().FirstOrDefault(ancestor => ancestor is W.Header or W.Footer or W.TextBoxContent)
            ?? element.Ancestors().LastOrDefault() ?? element;

        private static OpenXmlElement? NextNativeHeaderFooterFieldElement(OpenXmlElement element, OpenXmlElement story) {
            // Stay within this story and mirror the visible-run reader's revision exclusions.
            if (element is not (W.TextBoxContent or W.DeletedRun or W.MoveFromRun or W.InsertedRun or W.MoveToRun) && element.FirstChild != null)
                return element.FirstChild;
            for (OpenXmlElement? current = element; current != null && !ReferenceEquals(current, story); current = current.Parent) {
                if (current.NextSibling() is OpenXmlElement next) return next;
            }
            return null;
        }
    }
}
