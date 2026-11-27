namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        /// <summary>Retains a literal leading space; absent separators keep the established native authoring default.</summary>
        private static bool HasExplicitNoteReferenceSeparator(LegacyDocElement note) =>
            TryReadFirstNoteContentCharacter(note, out char character) && character == ' ';

        private static bool TryReadFirstNoteContentCharacter(LegacyDocElement element, out char character) {
            character = '\0';
            switch (element.WName) {
                case "t":
                case "delText":
                    string text = element.Text;
                    if (text.Length == 0) return false;
                    character = text[0];
                    return true;
                case "pPr":
                case "rPr":
                case "footnoteRef":
                case "endnoteRef":
                case "lastRenderedPageBreak":
                case "bookmarkStart":
                case "bookmarkEnd":
                    return false;
                case "sdt":
                    LegacyDocElement content = element.Child("sdtContent");
                    return content.IsNone || TryReadFirstNoteContentCharacter(content, out character);
                case "p":
                case "r":
                case "ins":
                case "del":
                case "hyperlink":
                case "footnote":
                case "endnote":
                case "sdtContent":
                    foreach (LegacyDocElement child in element.Elements()) {
                        if (TryReadFirstNoteContentCharacter(child, out character)) return true;
                    }
                    // An empty paragraph still ends before any following paragraph's text.
                    if (element.Is("p")) {
                        character = '\r';
                        return true;
                    }
                    return false;
                default:
                    // Field delimiters, breaks, pictures and unsupported elements are barriers, including empty elements.
                    return !IsIgnorableParagraphMarkup(element);
            }
        }
    }
}
