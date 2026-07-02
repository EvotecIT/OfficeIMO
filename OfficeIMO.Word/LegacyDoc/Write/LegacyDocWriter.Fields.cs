using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const string PageFieldInstruction = " PAGE   \\* MERGEFORMAT ";
        private const string NumberOfPagesFieldInstruction = " NUMPAGES   \\* MERGEFORMAT ";
        private const string DateFieldInstruction = " DATE ";
        private const string TimeFieldInstruction = " TIME ";
        private const string CreateDateFieldInstruction = " CREATEDATE ";
        private const string SaveDateFieldInstruction = " SAVEDATE ";
        private const string PrintDateFieldInstruction = " PRINTDATE ";
        private const string SupportedFieldNames = "PAGE, NUMPAGES, DATE, TIME, CREATEDATE, SAVEDATE, PRINTDATE, and document-property display fields";

        private static void AppendSupportedPageNumberField(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableFormatting formatting) {
            AppendSupportedField(text, runs, GetSupportedFieldInstruction(LegacyDocFieldKind.Page), "1", formatting);
        }

        private static void AppendSupportedField(StringBuilder text, List<LegacyDocWritableRun> runs, string instruction, string resultText, LegacyDocWritableFormatting formatting) {
            AppendFormattedText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, instruction, LegacyDocWritableFormatting.Plain);
            AppendFormattedText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, resultText, formatting);
            AppendFormattedText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
        }

        private static void AppendSupportedField(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            string instruction,
            string resultText,
            LegacyDocWritableFormatting formatting,
            IReadOnlyList<LegacyDocSimpleFieldBookmarkMarker> bookmarkMarkers,
            int characterOffset) {
            AppendFormattedText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, instruction, LegacyDocWritableFormatting.Plain);
            AppendFormattedText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            int resultStartCharacter = characterOffset + text.Length;
            string safeResultText = resultText;
            AppendFieldResultTextWithBookmarkMarkers(text, runs, bookmarks, safeResultText, formatting, bookmarkMarkers, resultStartCharacter);
            AppendFormattedText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
        }

        private static void AppendFieldResultTextWithBookmarkMarkers(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            string resultText,
            LegacyDocWritableFormatting formatting,
            IReadOnlyList<LegacyDocSimpleFieldBookmarkMarker> bookmarkMarkers,
            int resultStartCharacter) {
            int currentOffset = 0;
            int[] markerOffsets = bookmarkMarkers
                .Select(marker => marker.ResultOffset)
                .Where(offset => offset >= 0 && offset <= resultText.Length)
                .Distinct()
                .OrderBy(offset => offset)
                .ToArray();

            foreach (int markerOffset in markerOffsets) {
                if (markerOffset > currentOffset) {
                    AppendFormattedText(text, runs, resultText.Substring(currentOffset, markerOffset - currentOffset), formatting);
                    currentOffset = markerOffset;
                }

                AddSimpleFieldBookmarkMarkers(bookmarks, bookmarkMarkers, resultStartCharacter, markerOffset);
            }

            if (currentOffset < resultText.Length) {
                AppendFormattedText(text, runs, resultText.Substring(currentOffset), formatting);
            }
        }

        private static void AppendSupportedPageNumberFieldFromSimpleField(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableBookmarksBuilder bookmarks, SimpleField field, LegacyDocWritableFormatting inheritedFormatting) {
            if (!TryReadSupportedFieldKind(field.Instruction?.Value, out LegacyDocFieldKind fieldKind)) {
                throw new NotSupportedException($"Native DOC saving currently supports only {SupportedFieldNames} simple fields. Other field types are not supported yet.");
            }

            LegacyDocSimpleFieldResult result = ReadSimpleFieldResult(field);
            LegacyDocWritableFormatting formatting = result.Formatting
                .WithInheritedFormatting(inheritedFormatting);
            AppendSupportedField(
                text,
                runs,
                bookmarks,
                field.Instruction?.Value ?? GetSupportedFieldInstruction(fieldKind),
                GetSupportedFieldResultText(fieldKind, result.Text),
                formatting,
                result.BookmarkMarkers,
                characterOffset: 0);
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
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocWritableFormatting inheritedFormatting) {
            var instruction = new StringBuilder();
            var resultText = new StringBuilder();
            LegacyDocWritableFormatting? resultFormatting = null;
            var bookmarkMarkers = new List<LegacyDocSimpleFieldBookmarkMarker>();
            bool sawSeparator = false;
            int resultOffset = 0;
            int index = childIndex;
            for (; index < paragraphChildren.Count; index++) {
                OpenXmlElement fieldChild = paragraphChildren[index];
                if (fieldChild is BookmarkStart bookmarkStart && sawSeparator) {
                    bookmarkMarkers.Add(new LegacyDocSimpleFieldBookmarkMarker(bookmarkStart, null, resultOffset));
                    continue;
                }

                if (fieldChild is BookmarkEnd bookmarkEnd && sawSeparator) {
                    bookmarkMarkers.Add(new LegacyDocSimpleFieldBookmarkMarker(null, bookmarkEnd, resultOffset));
                    continue;
                }

                if (fieldChild is not Run run) {
                    if (IsIgnorableParagraphMarkup(fieldChild)) {
                        continue;
                    }

                    throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} complex fields only when the whole field is represented by adjacent runs.");
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
                        case Text textNode when sawSeparator:
                            resultFormatting ??= runFormatting;
                            if (!resultFormatting.Value.Equals(runFormatting)) {
                                throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} complex fields only when their display runs use one formatting set.");
                            }

                            resultText.Append(textNode.Text);
                            resultOffset = resultText.Length;
                            break;
                        case TabChar when sawSeparator:
                        case CarriageReturn when sawSeparator:
                        case NoBreakHyphen when sawSeparator:
                        case SoftHyphen when sawSeparator:
                        case Break when sawSeparator:
                            resultFormatting ??= runFormatting;
                            if (!resultFormatting.Value.Equals(runFormatting)) {
                                throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} complex fields only when their display runs use one formatting set.");
                            }

                            AppendSupportedFieldResultElementText(resultText, child);
                            resultOffset = resultText.Length;
                            break;
                        case FieldChar fieldChar:
                            FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                            if (fieldCharType == FieldCharValues.Begin) {
                                if (index != childIndex) {
                                    throw new NotSupportedException($"Native DOC saving does not support nested complex fields in {SupportedFieldNames} field runs.");
                                }

                                break;
                            }

                            if (fieldCharType == FieldCharValues.Separate) {
                                sawSeparator = true;
                                break;
                            }

                            if (fieldCharType == FieldCharValues.End) {
                                if (!TryReadSupportedFieldKind(instruction.ToString(), out LegacyDocFieldKind fieldKind)) {
                                    throw new NotSupportedException($"Native DOC saving currently supports only {SupportedFieldNames} complex fields. Other field types are not supported yet.");
                                }

                                LegacyDocWritableFormatting formatting = (resultFormatting ?? LegacyDocWritableFormatting.Plain)
                                    .WithInheritedFormatting(inheritedFormatting);
                                AppendSupportedField(
                                    text,
                                    runs,
                                    bookmarks,
                                    instruction.ToString(),
                                    GetSupportedFieldResultText(fieldKind, resultText.ToString()),
                                    formatting,
                                    bookmarkMarkers,
                                    characterOffset: 0);
                                childIndex = index;
                                return;
                            }

                            throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} complex fields only with begin, separate, and end field characters.");
                        default:
                            throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} complex fields only with field code and display text runs. Unsupported field run element: {child.LocalName}.");
                    }
                }
            }

            throw new NotSupportedException($"Native DOC saving cannot write an unterminated {SupportedFieldNames} complex field.");
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

            if (LegacyDocField.IsDocumentPropertyInstruction(trimmed)) {
                fieldKind = LegacyDocFieldKind.DocumentProperty;
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
                LegacyDocFieldKind.Date => DateFieldInstruction,
                LegacyDocFieldKind.Time => TimeFieldInstruction,
                LegacyDocFieldKind.CreateDate => CreateDateFieldInstruction,
                LegacyDocFieldKind.SaveDate => SaveDateFieldInstruction,
                LegacyDocFieldKind.PrintDate => PrintDateFieldInstruction,
                LegacyDocFieldKind.DocumentProperty => throw new NotSupportedException("Native DOC saving requires document-property fields to preserve their source instruction."),
                _ => throw new NotSupportedException($"Native DOC saving supports only {SupportedFieldNames} field instructions.")
            };
        }

        private static string GetSupportedFieldResultText(LegacyDocFieldKind fieldKind, string resultText) {
            return fieldKind == LegacyDocFieldKind.Page || fieldKind == LegacyDocFieldKind.NumPages
                ? "1"
                : resultText;
        }

        private static LegacyDocSimpleFieldResult ReadSimpleFieldResult(SimpleField field) {
            LegacyDocWritableFormatting? formatting = null;
            var bookmarkMarkers = new List<LegacyDocSimpleFieldBookmarkMarker>();
            var resultText = new StringBuilder();
            int resultOffset = 0;
            foreach (OpenXmlElement child in field.ChildElements) {
                switch (child) {
                    case Run run:
                        string runText = ReadSimpleFieldResultRunText(run);
                        LegacyDocWritableFormatting runFormatting = ReadSupportedRunFormatting(run.RunProperties);
                        formatting ??= runFormatting;
                        if (!formatting.Value.Equals(runFormatting)) {
                            throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} simple fields only when their display runs use one formatting set.");
                        }

                        resultText.Append(runText);
                        resultOffset = resultText.Length;
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarkMarkers.Add(new LegacyDocSimpleFieldBookmarkMarker(bookmarkStart, null, resultOffset));
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarkMarkers.Add(new LegacyDocSimpleFieldBookmarkMarker(null, bookmarkEnd, resultOffset));
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} simple fields only when their display result contains text runs and bookmarks. Unsupported field result element: {child.LocalName}.");
                }
            }

            return new LegacyDocSimpleFieldResult(resultText.ToString(), formatting ?? LegacyDocWritableFormatting.Plain, bookmarkMarkers);
        }

        private static string ReadSimpleFieldResultRunText(Run run) {
            var result = new StringBuilder();
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                    case LastRenderedPageBreak:
                        break;
                    case Text:
                    case TabChar:
                    case CarriageReturn:
                    case NoBreakHyphen:
                    case SoftHyphen:
                    case Break:
                        AppendSupportedFieldResultElementText(result, child);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} simple fields only when their display result contains text runs. Unsupported field result element: {child.LocalName}.");
                }
            }

            return result.ToString();
        }

        private static void AppendSupportedFieldResultElementText(StringBuilder result, OpenXmlElement element) {
            switch (element) {
                case Text text:
                    result.Append(text.Text);
                    break;
                case TabChar:
                    result.Append('\t');
                    break;
                case CarriageReturn:
                    result.Append(LegacyDocSpecialCharacters.TextWrappingBreak);
                    break;
                case NoBreakHyphen:
                    result.Append(LegacyDocSpecialCharacters.NoBreakHyphen);
                    break;
                case SoftHyphen:
                    result.Append(LegacyDocSpecialCharacters.SoftHyphen);
                    break;
                case Break breakNode:
                    result.Append(GetSupportedFieldResultBreakCharacter(breakNode));
                    break;
                default:
                    throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} field result runs only when they contain text, tabs, carriage returns, soft/no-break hyphens, and supported breaks. Unsupported field result element: {element.LocalName}.");
            }
        }

        private static char GetSupportedFieldResultBreakCharacter(Break breakNode) {
            BreakValues? breakType = breakNode.Type?.Value;
            if (breakType == null || breakType == BreakValues.TextWrapping) {
                return LegacyDocSpecialCharacters.TextWrappingBreak;
            }

            if (breakType == BreakValues.Page) {
                return LegacyDocSpecialCharacters.PageBreak;
            }

            if (breakType == BreakValues.Column) {
                return LegacyDocSpecialCharacters.ColumnBreak;
            }

            throw new NotSupportedException($"Native DOC saving currently supports text-wrapping, page, and column breaks in {SupportedFieldNames} field result runs only. Unsupported break type: {breakType}.");
        }

        private static void AddSimpleFieldBookmarkMarkers(
            LegacyDocWritableBookmarksBuilder bookmarks,
            IReadOnlyList<LegacyDocSimpleFieldBookmarkMarker> bookmarkMarkers,
            int resultStartCharacter,
            int resultOffset) {
            foreach (LegacyDocSimpleFieldBookmarkMarker marker in bookmarkMarkers) {
                if (marker.ResultOffset != resultOffset) {
                    continue;
                }

                int characterPosition = resultStartCharacter + resultOffset;
                if (marker.BookmarkStart != null) {
                    bookmarks.AddStart(marker.BookmarkStart, characterPosition);
                } else if (marker.BookmarkEnd != null) {
                    bookmarks.AddEnd(marker.BookmarkEnd, characterPosition);
                }
            }
        }

        private readonly struct LegacyDocSimpleFieldResult {
            internal LegacyDocSimpleFieldResult(string text, LegacyDocWritableFormatting formatting, IReadOnlyList<LegacyDocSimpleFieldBookmarkMarker> bookmarkMarkers) {
                Text = text;
                Formatting = formatting;
                BookmarkMarkers = bookmarkMarkers;
            }

            internal string Text { get; }

            internal LegacyDocWritableFormatting Formatting { get; }

            internal IReadOnlyList<LegacyDocSimpleFieldBookmarkMarker> BookmarkMarkers { get; }
        }

        private readonly struct LegacyDocSimpleFieldBookmarkMarker {
            internal LegacyDocSimpleFieldBookmarkMarker(BookmarkStart? bookmarkStart, BookmarkEnd? bookmarkEnd, int resultOffset) {
                BookmarkStart = bookmarkStart;
                BookmarkEnd = bookmarkEnd;
                ResultOffset = resultOffset;
            }

            internal BookmarkStart? BookmarkStart { get; }

            internal BookmarkEnd? BookmarkEnd { get; }

            internal int ResultOffset { get; }
        }
    }
}
