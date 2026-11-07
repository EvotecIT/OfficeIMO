using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        /// <summary>Retains a literal leading space; absent separators keep the established native authoring default.</summary>
        private static bool HasExplicitNoteReferenceSeparator(OpenXmlElement note) =>
            TryReadFirstNoteContentCharacter(note, out char character) && character == ' ';

        private static bool TryReadFirstNoteContentCharacter(OpenXmlElement element, out char character) {
            character = '\0';
            switch (element) {
                case Text text:
                    if (text.Text.Length == 0) return false;
                    character = text.Text[0];
                    return true;
                case DeletedText text:
                    if (text.Text.Length == 0) return false;
                    character = text.Text[0];
                    return true;
                case ParagraphProperties:
                case RunProperties:
                case FootnoteReferenceMark:
                case EndnoteReferenceMark:
                case LastRenderedPageBreak:
                case BookmarkStart:
                case BookmarkEnd:
                    return false;
                case SdtRun control:
                    return control.SdtContentRun == null || TryReadFirstNoteContentCharacter(control.SdtContentRun, out character);
                case SdtBlock control:
                    return control.SdtContentBlock == null || TryReadFirstNoteContentCharacter(control.SdtContentBlock, out character);
                case Paragraph:
                case Run:
                case InsertedRun:
                case DeletedRun:
                case Hyperlink:
                case Footnote:
                case Endnote:
                case SdtContentRun:
                case SdtContentBlock:
                    foreach (OpenXmlElement child in element.ChildElements) {
                        if (TryReadFirstNoteContentCharacter(child, out character)) return true;
                    }
                    // An empty paragraph still ends before any following paragraph's text.
                    if (element is Paragraph) {
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
