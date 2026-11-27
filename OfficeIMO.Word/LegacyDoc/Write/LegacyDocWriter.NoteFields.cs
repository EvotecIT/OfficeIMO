using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedNoteFieldFromSimpleField(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocElement field,
            int storyStart,
            bool allowHyperlinkRunStyle = false) {
            string? fieldInstruction = field.Attr("instr");
            if (!TryReadSupportedFieldKind(fieldInstruction, out LegacyDocFieldKind fieldKind)) {
                throw new NotSupportedException($"Native DOC saving currently supports only {SupportedFieldNames} simple fields in note paragraphs. Other field types are not supported yet.");
            }
            LegacyDocWritableFieldResult result = ReadSimpleFieldResult(field, allowHyperlinkRunStyle);
            AppendSupportedField(text, runs, bookmarks, fieldInstruction ?? GetSupportedFieldInstruction(fieldKind), result, GetSupportedFieldEmptyResultText(fieldKind), storyStart);
        }

        private static void AppendSupportedNoteComplexPageNumberField(
            IReadOnlyList<LegacyDocElement> paragraphChildren,
            ref int childIndex,
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            int storyStart) {
            LegacyDocWritableFieldResult result = ReadSupportedComplexField(paragraphChildren, ref childIndex,
                LegacyDocWritableFormatting.Plain, out string instruction, out LegacyDocFieldKind fieldKind);
            AppendSupportedField(text, runs, bookmarks, instruction, result, GetSupportedFieldEmptyResultText(fieldKind), storyStart);
        }
    }
}
