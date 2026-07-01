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

        private static bool IsComplexFieldBeginRun(Run run) {
            FieldChar? fieldChar = run.Elements<FieldChar>().FirstOrDefault();
            return fieldChar?.FieldCharType?.Value == FieldCharValues.Begin;
        }

        private static void AppendSupportedComplexPageNumberField(
            IReadOnlyList<OpenXmlElement> paragraphChildren,
            ref int childIndex,
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableFormatting inheritedFormatting) {
            var instruction = new StringBuilder();
            LegacyDocWritableFormatting? resultFormatting = null;
            bool sawSeparator = false;
            int index = childIndex;
            for (; index < paragraphChildren.Count; index++) {
                if (paragraphChildren[index] is not Run run) {
                    if (IsIgnorableParagraphMarkup(paragraphChildren[index])) {
                        continue;
                    }

                    throw new NotSupportedException("Native DOC saving supports PAGE and NUMPAGES complex fields only when the whole field is represented by adjacent runs.");
                }

                LegacyDocWritableFormatting runFormatting = ReadSupportedRunFormatting(run.RunProperties);
                foreach (OpenXmlElement child in run.ChildElements) {
                    switch (child) {
                        case RunProperties:
                        case LastRenderedPageBreak:
                            break;
                        case FieldCode fieldCode when !sawSeparator:
                            instruction.Append(fieldCode.Text);
                            break;
                        case Text when sawSeparator:
                            resultFormatting ??= runFormatting;
                            if (!resultFormatting.Value.Equals(runFormatting)) {
                                throw new NotSupportedException("Native DOC saving supports PAGE and NUMPAGES complex fields only when their display runs use one formatting set.");
                            }

                            break;
                        case FieldChar fieldChar:
                            FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                            if (fieldCharType == FieldCharValues.Begin) {
                                if (index != childIndex) {
                                    throw new NotSupportedException("Native DOC saving does not support nested complex fields in PAGE and NUMPAGES field runs.");
                                }

                                break;
                            }

                            if (fieldCharType == FieldCharValues.Separate) {
                                sawSeparator = true;
                                break;
                            }

                            if (fieldCharType == FieldCharValues.End) {
                                if (!TryReadSupportedFieldKind(instruction.ToString(), out LegacyDocFieldKind fieldKind)) {
                                    throw new NotSupportedException("Native DOC saving currently supports only PAGE and NUMPAGES complex fields. Other field types are not supported yet.");
                                }

                                LegacyDocWritableFormatting formatting = (resultFormatting ?? LegacyDocWritableFormatting.Plain)
                                    .WithInheritedFormatting(inheritedFormatting);
                                AppendSupportedField(text, runs, fieldKind, formatting);
                                childIndex = index;
                                return;
                            }

                            throw new NotSupportedException("Native DOC saving supports PAGE and NUMPAGES complex fields only with begin, separate, and end field characters.");
                        default:
                            throw new NotSupportedException($"Native DOC saving supports PAGE and NUMPAGES complex fields only with field code and display text runs. Unsupported field run element: {child.LocalName}.");
                    }
                }
            }

            throw new NotSupportedException("Native DOC saving cannot write an unterminated PAGE or NUMPAGES complex field.");
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
                    throw new NotSupportedException("Native DOC saving supports PAGE and NUMPAGES simple fields only when their display runs use one formatting set.");
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
