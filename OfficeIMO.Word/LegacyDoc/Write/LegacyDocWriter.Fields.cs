using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const string PageFieldInstruction = " PAGE   \\* MERGEFORMAT ";
        private const string NumberOfPagesFieldInstruction = " NUMPAGES   \\* MERGEFORMAT ";
        private const string SectionPagesFieldInstruction = " SECTIONPAGES   \\* MERGEFORMAT ";
        private const string DateFieldInstruction = " DATE ";
        private const string TimeFieldInstruction = " TIME ";
        private const string CreateDateFieldInstruction = " CREATEDATE ";
        private const string SaveDateFieldInstruction = " SAVEDATE ";
        private const string PrintDateFieldInstruction = " PRINTDATE ";
        private const string SupportedFieldNames = "PAGE, NUMPAGES, SECTIONPAGES, DATE, TIME, CREATEDATE, SAVEDATE, PRINTDATE, FILENAME, STYLEREF, EQ, and document-property display fields";

        private static void AppendSupportedPageNumberField(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableFormatting formatting) {
            AppendSupportedField(text, runs, GetSupportedFieldInstruction(LegacyDocFieldKind.Page), "1", formatting);
        }

        private static void AppendMathEquationField(StringBuilder text, List<LegacyDocWritableRun> runs, OpenXmlElement mathElement, LegacyDocWritableFormatting formatting) {
            AppendSupportedField(
                text,
                runs,
                WordMath.ToEquationFieldInstruction(mathElement),
                WordMath.GetText(mathElement),
                formatting);
        }

        private static void AppendMathEquationNoteField(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            OpenXmlElement mathElement,
            int storyStart) {
            var result = new LegacyDocWritableFieldResult();
            result.AppendText(WordMath.GetText(mathElement), LegacyDocWritableFormatting.Plain);
            AppendSupportedField(text, runs, bookmarks, WordMath.ToEquationFieldInstruction(mathElement), result, GetSupportedFieldEmptyResultText(LegacyDocFieldKind.Equation), storyStart);
        }

        private static void AppendSupportedField(StringBuilder text, List<LegacyDocWritableRun> runs, string instruction, string resultText, LegacyDocWritableFormatting formatting) {
            AppendFormattedText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, instruction, LegacyDocWritableFormatting.Plain);
            AppendFormattedText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, resultText, formatting);
            AppendFormattedText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
        }

        private static void AppendSupportedPageNumberFieldFromSimpleField(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            SimpleField field,
            LegacyDocWritableFormatting inheritedFormatting,
            bool allowHyperlinkRunStyle = false) {
            if (!TryReadSupportedFieldKind(field.Instruction?.Value, out LegacyDocFieldKind fieldKind)) {
                throw new NotSupportedException($"Native DOC saving currently supports only {SupportedFieldNames} simple fields. Other field types are not supported yet.");
            }

            LegacyDocWritableFieldResult result = ReadSimpleFieldResult(field, allowHyperlinkRunStyle, inheritedFormatting);
            AppendSupportedField(
                text,
                runs,
                bookmarks,
                field.Instruction?.Value ?? GetSupportedFieldInstruction(fieldKind),
                result,
                GetSupportedFieldEmptyResultText(fieldKind),
                characterOffset: 0);
        }

        private static bool IsComplexFieldBeginRun(Run run) {
            FieldChar? fieldChar = run.Elements<FieldChar>().FirstOrDefault();
            return fieldChar?.FieldCharType?.Value == FieldCharValues.Begin;
        }

        private static bool TryReadSupportedFieldKind(string? instruction, out LegacyDocFieldKind fieldKind) {
            fieldKind = LegacyDocFieldKind.None;
            if (string.IsNullOrWhiteSpace(instruction)) {
                return false;
            }

            string trimmed = instruction!.Trim();
            if (IsFieldInstruction(trimmed, "PAGE")) {
                fieldKind = LegacyDocFieldKind.Page;
                return true;
            }

            if (IsFieldInstruction(trimmed, "NUMPAGES")) {
                fieldKind = LegacyDocFieldKind.NumPages;
                return true;
            }

            if (IsFieldInstruction(trimmed, "SECTIONPAGES")) {
                fieldKind = LegacyDocFieldKind.SectionPages;
                return true;
            }

            if (IsFieldInstruction(trimmed, "DATE")) {
                fieldKind = LegacyDocFieldKind.Date;
                return true;
            }

            if (IsFieldInstruction(trimmed, "TIME")) {
                fieldKind = LegacyDocFieldKind.Time;
                return true;
            }

            if (IsFieldInstruction(trimmed, "CREATEDATE")) {
                fieldKind = LegacyDocFieldKind.CreateDate;
                return true;
            }

            if (IsFieldInstruction(trimmed, "SAVEDATE")) {
                fieldKind = LegacyDocFieldKind.SaveDate;
                return true;
            }

            if (IsFieldInstruction(trimmed, "PRINTDATE")) {
                fieldKind = LegacyDocFieldKind.PrintDate;
                return true;
            }

            if (LegacyDocField.IsTextReferenceInstruction(trimmed)) {
                fieldKind = LegacyDocFieldKind.TextReference;
                return true;
            }

            if (LegacyDocField.IsDocumentPropertyInstruction(trimmed)) {
                fieldKind = LegacyDocFieldKind.DocumentProperty;
                return true;
            }

            if (IsFieldInstruction(trimmed, "EQ")) {
                fieldKind = LegacyDocFieldKind.Equation;
                return true;
            }

            return false;
        }

        private static bool IsFieldInstruction(string trimmedInstruction, string fieldName) {
            return trimmedInstruction.StartsWith(fieldName, StringComparison.OrdinalIgnoreCase)
                && (trimmedInstruction.Length == fieldName.Length || char.IsWhiteSpace(trimmedInstruction[fieldName.Length]));
        }

        private static string GetSupportedFieldInstruction(LegacyDocFieldKind fieldKind) {
            return fieldKind switch {
                LegacyDocFieldKind.Page => PageFieldInstruction,
                LegacyDocFieldKind.NumPages => NumberOfPagesFieldInstruction,
                LegacyDocFieldKind.SectionPages => SectionPagesFieldInstruction,
                LegacyDocFieldKind.Date => DateFieldInstruction,
                LegacyDocFieldKind.Time => TimeFieldInstruction,
                LegacyDocFieldKind.CreateDate => CreateDateFieldInstruction,
                LegacyDocFieldKind.SaveDate => SaveDateFieldInstruction,
                LegacyDocFieldKind.PrintDate => PrintDateFieldInstruction,
                LegacyDocFieldKind.DocumentProperty => throw new NotSupportedException("Native DOC saving requires document-property fields to preserve their source instruction."),
                LegacyDocFieldKind.TextReference => throw new NotSupportedException("Native DOC saving requires text-reference fields to preserve their source instruction."),
                LegacyDocFieldKind.Equation => " EQ ",
                _ => throw new NotSupportedException($"Native DOC saving supports only {SupportedFieldNames} field instructions.")
            };
        }

        private static string GetSupportedFieldEmptyResultText(LegacyDocFieldKind fieldKind) {
            return fieldKind == LegacyDocFieldKind.Page || fieldKind == LegacyDocFieldKind.NumPages || fieldKind == LegacyDocFieldKind.SectionPages
                ? "1"
                : string.Empty;
        }
    }
}
