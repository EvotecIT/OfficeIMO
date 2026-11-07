using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedNoteFieldFromSimpleField(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            SimpleField field,
            int storyStart,
            bool allowHyperlinkRunStyle = false) {
            if (!TryReadSupportedFieldKind(field.Instruction?.Value, out LegacyDocFieldKind fieldKind)) {
                throw new NotSupportedException($"Native DOC saving currently supports only {SupportedFieldNames} simple fields in note paragraphs. Other field types are not supported yet.");
            }
            LegacyDocWritableFieldResult result = ReadSimpleFieldResult(field, allowHyperlinkRunStyle);
            AppendSupportedField(text, runs, bookmarks, field.Instruction?.Value ?? GetSupportedFieldInstruction(fieldKind), result, GetSupportedFieldEmptyResultText(fieldKind), storyStart);
        }

        private static void AppendSupportedNoteComplexPageNumberField(
            IReadOnlyList<OpenXmlElement> paragraphChildren,
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
