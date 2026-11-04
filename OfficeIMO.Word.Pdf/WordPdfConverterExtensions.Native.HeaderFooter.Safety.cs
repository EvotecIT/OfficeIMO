using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const int MaximumNativeHeaderFooterTextBoxNestingDepth = 16;

        private static void ValidateNativeHeaderFooterTextBoxNesting(WordHeaderFooter headerFooter) {
            OpenXmlElement? root = (OpenXmlElement?)headerFooter._header ?? headerFooter._footer;
            if (root == null) {
                return;
            }

            var pending = new Stack<(OpenXmlElement Element, int TextBoxDepth)>();
            for (OpenXmlElement? child = root.LastChild; child != null; child = child.PreviousSibling()) {
                pending.Push((child, 0));
            }

            while (pending.Count > 0) {
                (OpenXmlElement element, int textBoxDepth) = pending.Pop();
                int nextDepth = element is W.TextBoxContent ? textBoxDepth + 1 : textBoxDepth;
                if (nextDepth > MaximumNativeHeaderFooterTextBoxNestingDepth) {
                    throw new InvalidDataException(
                        $"Header or footer text-box nesting exceeds the supported limit of {MaximumNativeHeaderFooterTextBoxNestingDepth} levels.");
                }

                for (OpenXmlElement? child = element.LastChild; child != null; child = child.PreviousSibling()) {
                    pending.Push((child, nextDepth));
                }
            }
        }
    }
}
