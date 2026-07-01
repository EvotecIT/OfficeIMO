using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const string PageFieldInstruction = " PAGE   \\* MERGEFORMAT ";
        private const string NumberOfPagesFieldInstruction = " NUMPAGES   \\* MERGEFORMAT ";

        private static void AppendSupportedPageNumberField(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableFormatting formatting) {
            AppendSupportedField(text, runs, LegacyDocFieldKind.Page, formatting);
        }

        private static void AppendSupportedField(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocFieldKind fieldKind, LegacyDocWritableFormatting formatting) {
            AppendFormattedText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, GetSupportedFieldInstruction(fieldKind), LegacyDocWritableFormatting.Plain);
            AppendFormattedText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, "1", formatting);
            AppendFormattedText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
        }

        private static void AppendSupportedPageNumberFieldFromSimpleField(StringBuilder text, List<LegacyDocWritableRun> runs, SimpleField field, LegacyDocWritableFormatting inheritedFormatting) {
            if (!TryReadSupportedFieldKind(field.Instruction?.Value, out LegacyDocFieldKind fieldKind)) {
                throw new NotSupportedException("Native DOC saving currently supports only PAGE and NUMPAGES simple fields. Other field types are not supported yet.");
            }

            LegacyDocWritableFormatting formatting = ReadSimpleFieldResultFormatting(field)
                .WithInheritedFormatting(inheritedFormatting);
            AppendSupportedField(text, runs, fieldKind, formatting);
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
                _ => throw new NotSupportedException("Native DOC saving supports only PAGE and NUMPAGES field instructions.")
            };
        }

        private static LegacyDocWritableFormatting ReadSimpleFieldResultFormatting(SimpleField field) {
            LegacyDocWritableFormatting? formatting = null;
            foreach (Run run in field.Elements<Run>()) {
                EnsureSimplePageFieldResultRun(run);
                LegacyDocWritableFormatting runFormatting = ReadSupportedRunFormatting(run.RunProperties);
                formatting ??= runFormatting;
                if (!formatting.Value.Equals(runFormatting)) {
                    throw new NotSupportedException("Native DOC saving supports PAGE simple fields only when their display runs use one formatting set.");
                }
            }

            return formatting ?? LegacyDocWritableFormatting.Plain;
        }

        private static void EnsureSimplePageFieldResultRun(Run run) {
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                    case LastRenderedPageBreak:
                    case Text:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports PAGE and NUMPAGES simple fields only when their display result contains text runs. Unsupported field result element: {child.LocalName}.");
                }
            }
        }
    }
}
