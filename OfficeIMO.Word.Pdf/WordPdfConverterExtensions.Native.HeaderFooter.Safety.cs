
namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private const int MaximumNativeHeaderFooterTextBoxNestingDepth = 16;

        private static void ValidateNativeHeaderFooterTextBoxNesting(WordHeaderFooter headerFooter) {
            OfficeXmlNode root = headerFooter.Node();
            if (root.IsNone) {
                return;
            }

            var pending = new Stack<(OfficeXmlNode Element, int TextBoxDepth)>();
            for (OfficeXmlNode child = root.LastChild; child.Exists; child = child.PreviousSibling) {
                pending.Push((child, 0));
            }

            while (pending.Count > 0) {
                (OfficeXmlNode element, int textBoxDepth) = pending.Pop();
                int nextDepth = element.IsW("txbxContent") ? textBoxDepth + 1 : textBoxDepth;
                if (nextDepth > MaximumNativeHeaderFooterTextBoxNestingDepth) {
                    throw new InvalidDataException(
                        $"Header or footer text-box nesting exceeds the supported limit of {MaximumNativeHeaderFooterTextBoxNestingDepth} levels.");
                }

                for (OfficeXmlNode child = element.LastChild; child.Exists; child = child.PreviousSibling) {
                    pending.Push((child, nextDepth));
                }
            }
        }
    }
}
