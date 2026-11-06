using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedComplexPageNumberField(
            IReadOnlyList<OpenXmlElement> paragraphChildren,
            ref int childIndex,
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocWritableFormatting inheritedFormatting) {
            LegacyDocWritableFieldResult result = ReadSupportedComplexField(paragraphChildren, ref childIndex,
                inheritedFormatting, out string instruction, out LegacyDocFieldKind fieldKind);
            AppendSupportedField(text, runs, bookmarks, instruction, result, GetSupportedFieldEmptyResultText(fieldKind), characterOffset: 0);
        }

        /// <summary>Reads one adjacent-run field; result content uses the same collector as simple fields and note stories.</summary>
        private static LegacyDocWritableFieldResult ReadSupportedComplexField(
            IReadOnlyList<OpenXmlElement> paragraphChildren,
            ref int childIndex,
            LegacyDocWritableFormatting inheritedFormatting,
            out string fieldInstruction,
            out LegacyDocFieldKind fieldKind) {
            var instruction = new StringBuilder();
            var result = new LegacyDocWritableFieldResult(inheritedFormatting);
            bool sawBegin = false;
            bool sawSeparator = false;
            for (int index = childIndex; index < paragraphChildren.Count; index++) {
                OpenXmlElement fieldChild = paragraphChildren[index];
                if (fieldChild is not Run run) {
                    if (sawSeparator) {
                        result.AppendChild(fieldChild);
                    } else if (!IsIgnorableParagraphMarkup(fieldChild)) {
                        throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} complex field instructions only when represented by adjacent runs.");
                    }
                    continue;
                }

                LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.RunProperties)
                    .WithInheritedFormatting(inheritedFormatting);
                foreach (OpenXmlElement child in run.ChildElements) {
                    switch (child) {
                        case RunProperties:
                        case LastRenderedPageBreak:
                            break;
                        case FieldCode code when !sawSeparator:
                            instruction.Append(code.Text);
                            break;
                        case FieldChar fieldChar:
                            FieldCharValues? kind = fieldChar.FieldCharType?.Value;
                            if (fieldChar.Dirty?.Value == true) {
                                throw new NotSupportedException("Native DOC saving cannot preserve an invalidated field result (w:dirty=true) without a qualified invalidation mapping.");
                            }
                            if (kind != FieldCharValues.Begin && fieldChar.FieldLock?.Value == true) {
                                throw new NotSupportedException("Native DOC saving supports complex field locks on the begin field character only.");
                            }
                            if (kind == FieldCharValues.Begin) {
                                if (sawBegin || index != childIndex) {
                                    throw new NotSupportedException($"Native DOC saving does not support nested complex fields in {SupportedFieldNames} field runs.");
                                }
                                result.FieldLocked = fieldChar.FieldLock?.Value == true;
                                sawBegin = true;
                            } else if (kind == FieldCharValues.Separate && !sawSeparator) {
                                sawSeparator = true;
                            } else if (kind == FieldCharValues.End) {
                                fieldInstruction = instruction.ToString();
                                if (!TryReadSupportedFieldKind(fieldInstruction, out fieldKind)) {
                                    throw new NotSupportedException($"Native DOC saving currently supports only {SupportedFieldNames} complex fields. Other field types are not supported yet.");
                                }
                                childIndex = index;
                                return result;
                            } else {
                                throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} complex fields only with begin, separate, and end field characters and at most one separator.");
                            }
                            break;
                        default:
                            if (!sawSeparator) {
                                throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} complex fields only with field code and display text runs. Unsupported field run element: {child.LocalName}.");
                            }
                            result.AppendElement(child, formatting);
                            break;
                    }
                }
            }
            throw new NotSupportedException($"Native DOC saving cannot write an unterminated {SupportedFieldNames} complex field.");
        }
    }
}
