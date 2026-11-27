using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedComplexPageNumberField(
            IReadOnlyList<LegacyDocElement> paragraphChildren,
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
            IReadOnlyList<LegacyDocElement> paragraphChildren,
            ref int childIndex,
            LegacyDocWritableFormatting inheritedFormatting,
            out string fieldInstruction,
            out LegacyDocFieldKind fieldKind) {
            var instruction = new StringBuilder();
            var result = new LegacyDocWritableFieldResult(inheritedFormatting);
            bool sawBegin = false;
            bool sawSeparator = false;
            for (int index = childIndex; index < paragraphChildren.Count; index++) {
                LegacyDocElement fieldChild = paragraphChildren[index];
                if (!fieldChild.Is("r")) {
                    if (sawSeparator) {
                        result.AppendChild(fieldChild);
                    } else if (!IsIgnorableParagraphMarkup(fieldChild)) {
                        throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} complex field instructions only when represented by adjacent runs.");
                    }
                    continue;
                }

                LegacyDocElement run = fieldChild;
                LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.Child("rPr"))
                    .WithInheritedFormatting(inheritedFormatting);
                foreach (LegacyDocElement child in run.Elements()) {
                    switch (child.WName) {
                        case "rPr":
                        case "lastRenderedPageBreak":
                            break;
                        case "instrText" when !sawSeparator:
                            instruction.Append(child.Text);
                            break;
                        case "fldChar":
                            string? kind = child.Attr("fldCharType");
                            if (kind == "begin") {
                                if (sawBegin || index != childIndex) {
                                    throw new NotSupportedException($"Native DOC saving does not support nested complex fields in {SupportedFieldNames} field runs.");
                                }
                                // The begin character owns the lock. Public setters also repeat it on separator/end characters.
                                result.FieldLocked = child.Attr("fldLock") is { } fieldLock && LegacyDocValues.ParseOnOff(fieldLock);
                                sawBegin = true;
                            } else if (kind == "separate" && !sawSeparator) {
                                sawSeparator = true;
                            } else if (kind == "end") {
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
