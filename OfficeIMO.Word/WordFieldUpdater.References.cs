using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private static bool TryEvaluateRef(
            WordDocument document,
            MutableFieldCandidate candidate,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;

            string? bookmarkName = GetBookmarkInstruction(parsed);
            if (string.IsNullOrWhiteSpace(bookmarkName)) {
                message = "REF field is missing a bookmark name.";
                return false;
            }

            if (!TryGetReferenceSwitches(parsed, out ReferenceListSwitch? listSwitch, out string? unsupportedReferenceSwitch, out string? switchError)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = switchError ?? $"REF field switch {unsupportedReferenceSwitch} is not evaluated by OfficeIMO.";
                return false;
            }

            BookmarkStart? bookmarkStart = FindBookmarkStart(document, bookmarkName!);
            if (bookmarkStart == null) {
                message = $"Bookmark {bookmarkName} was not found.";
                return false;
            }

            if (listSwitch.HasValue) {
                if (!TryResolveReferenceListNumber(document, candidate, bookmarkStart, listSwitch.Value, parsed.Switches, out string? listNumber, out WordFieldUpdateStatus listStatus, out string listMessage)) {
                    status = listStatus;
                    message = listMessage;
                    return false;
                }

                value = listNumber;
                status = WordFieldUpdateStatus.Updated;
                message = $"Updated paragraph number for bookmark {bookmarkName} using REF {ReferenceListSwitchToFieldCode(listSwitch.Value)}.";
                return true;
            }

            string bookmarkText = GetBookmarkRangeText(bookmarkStart);
            if (string.IsNullOrEmpty(bookmarkText)) {
                Paragraph? paragraph = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault();
                bookmarkText = paragraph == null ? string.Empty : GetVisibleParagraphText(paragraph);
            }

            if (string.IsNullOrEmpty(bookmarkText)) {
                message = $"Bookmark {bookmarkName} did not contain visible text.";
                return false;
            }

            if (!TryApplyReferenceTextFormat(parsed.FormatSwitches, bookmarkText, out string formattedText, out string? unsupportedFormat)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = $"REF field format switch {unsupportedFormat} is not supported for bookmark text refresh.";
                return false;
            }

            value = formattedText;
            status = WordFieldUpdateStatus.Updated;
            message = parsed.FormatSwitches.Any(fieldFormat => fieldFormat != WordFieldFormat.Mergeformat && fieldFormat != WordFieldFormat.CharFormat)
                ? $"Updated from bookmark {bookmarkName} with deterministic text formatting."
                : $"Updated from bookmark {bookmarkName}.";
            return true;
        }

        private static bool TryEvaluatePageRef(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;

            string? bookmarkName = GetBookmarkInstruction(parsed);
            if (string.IsNullOrWhiteSpace(bookmarkName)) {
                message = "PAGEREF field is missing a bookmark name.";
                return false;
            }

            if (HasUnsupportedPageReferenceSwitches(parsed)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = "PAGEREF field uses switches that need Word layout context and were left unchanged.";
                return false;
            }

            BookmarkStart? bookmarkStart = FindBookmarkStart(document, bookmarkName!);
            if (bookmarkStart == null) {
                message = $"Bookmark {bookmarkName} was not found.";
                return false;
            }

            if (!bookmarkStart.Ancestors<Body>().Any()) {
                message = "PAGEREF fields outside the document body need Word layout context and were left unchanged.";
                return false;
            }

            int? page = EstimatePageForBodyField(document, bookmarkStart);
            if (page == null) {
                message = $"Bookmark {bookmarkName} could not be matched to a body paragraph.";
                return false;
            }

            if (!TryFormatPageReferenceValue(page.Value, parsed.FormatSwitches, parsed.NumericPictureSwitch, out string pageText, out string? unsupportedFormat, out bool usedNumericPicture)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = $"PAGEREF field format switch {unsupportedFormat} is not supported for deterministic page-reference refresh.";
                return false;
            }

            value = pageText;
            status = WordFieldUpdateStatus.Updated;
            message = usedNumericPicture
                ? $"Updated from OfficeIMO page-break order for bookmark {bookmarkName} with numeric picture formatting."
                : parsed.FormatSwitches.Any(fieldFormat => fieldFormat != WordFieldFormat.Mergeformat && fieldFormat != WordFieldFormat.CharFormat)
                ? $"Updated from OfficeIMO page-break order for bookmark {bookmarkName} with deterministic number formatting."
                : $"Updated from OfficeIMO page-break order for bookmark {bookmarkName}.";
            return true;
        }

        private static string? GetBookmarkInstruction(WordFieldInventory.ParsedFieldInstruction parsed) {
            string? bookmarkName = parsed.Instructions.FirstOrDefault();
            return string.IsNullOrWhiteSpace(bookmarkName) ? null : TrimQuotes(bookmarkName);
        }

        private static bool HasUnsupportedPageReferenceSwitches(WordFieldInventory.ParsedFieldInstruction parsed) {
            return parsed.Switches.Any(fieldSwitch =>
                !string.Equals(fieldSwitch, "\\h", StringComparison.OrdinalIgnoreCase));
        }

        private static bool TryGetReferenceSwitches(
            WordFieldInventory.ParsedFieldInstruction parsed,
            out ReferenceListSwitch? listSwitch,
            out string? unsupportedSwitch,
            out string? error) {
            listSwitch = null;
            unsupportedSwitch = null;
            error = null;

            foreach (string fieldSwitch in parsed.Switches) {
                string trimmed = fieldSwitch.Trim();
                if (string.Equals(trimmed, "\\h", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                ReferenceListSwitch? current = null;
                if (string.Equals(trimmed, "\\n", StringComparison.OrdinalIgnoreCase)) {
                    current = ReferenceListSwitch.NoContext;
                } else if (string.Equals(trimmed, "\\r", StringComparison.OrdinalIgnoreCase)) {
                    current = ReferenceListSwitch.RelativeContext;
                } else if (string.Equals(trimmed, "\\w", StringComparison.OrdinalIgnoreCase)) {
                    current = ReferenceListSwitch.FullContext;
                } else if (string.Equals(trimmed, "\\t", StringComparison.OrdinalIgnoreCase) && HasReferenceListSwitch(parsed.Switches)) {
                    continue;
                }

                if (current.HasValue) {
                    if (listSwitch.HasValue) {
                        error = "REF field cannot combine multiple paragraph-number switches.";
                        return false;
                    }

                    listSwitch = current.Value;
                    continue;
                }

                unsupportedSwitch = trimmed;
                return false;
            }

            return true;
        }

        private static bool HasReferenceListSwitch(IReadOnlyList<string> switches) {
            return switches.Any(fieldSwitch =>
                string.Equals(fieldSwitch.Trim(), "\\n", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(fieldSwitch.Trim(), "\\r", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(fieldSwitch.Trim(), "\\w", StringComparison.OrdinalIgnoreCase));
        }

        private static bool TryApplyReferenceTextFormat(
            IReadOnlyList<WordFieldFormat> formatSwitches,
            string source,
            out string value,
            out string? unsupportedFormat) {
            value = source;
            unsupportedFormat = null;

            WordFieldFormat? format = GetLastMeaningfulFormat(formatSwitches);
            switch (format) {
                case null:
                    return true;
                case WordFieldFormat.Lower:
                    value = source.ToLowerInvariant();
                    return true;
                case WordFieldFormat.Upper:
                    value = source.ToUpperInvariant();
                    return true;
                case WordFieldFormat.FirstCap:
                    value = CapitalizeFirstWord(source);
                    return true;
                case WordFieldFormat.Caps:
                    value = CapitalizeEachWord(source);
                    return true;
                default:
                    unsupportedFormat = format.Value.ToString();
                    return false;
            }
        }

        private static bool TryFormatPageReferenceValue(
            int page,
            IReadOnlyList<WordFieldFormat> formatSwitches,
            string? numericPicture,
            out string value,
            out string? unsupportedFormat,
            out bool usedNumericPicture) {
            unsupportedFormat = null;
            usedNumericPicture = false;
            if (!string.IsNullOrWhiteSpace(numericPicture)) {
                usedNumericPicture = true;
                if (GetLastMeaningfulFormat(formatSwitches) != null) {
                    value = string.Empty;
                    unsupportedFormat = @"combined \# numeric picture and \* format switches";
                    return false;
                }

                if (!TryFormatFormulaValue(page, numericPicture, out value, out string? diagnostic)) {
                    unsupportedFormat = diagnostic == null
                        ? @"\# numeric picture"
                        : diagnostic.Replace("Formula numeric picture", "Field numeric picture");
                    return false;
                }

                return true;
            }

            WordFieldFormat? format = GetLastMeaningfulFormat(formatSwitches);

            switch (format) {
                case null:
                case WordFieldFormat.Arabic:
                    value = page.ToString(CultureInfo.InvariantCulture);
                    return true;
                case WordFieldFormat.Roman:
                case WordFieldFormat.roman:
                case WordFieldFormat.Ordinal:
                case WordFieldFormat.Alphabetical:
                case WordFieldFormat.ALPHABETICAL:
                case WordFieldFormat.Hex:
                case WordFieldFormat.CardText:
                case WordFieldFormat.OrdText:
                case WordFieldFormat.DollarText:
                    value = FormatSequenceValue(page, new[] { format.Value });
                    return true;
                default:
                    value = string.Empty;
                    unsupportedFormat = format.Value.ToString();
                    return false;
            }
        }

        private static WordFieldFormat? GetLastMeaningfulFormat(IReadOnlyList<WordFieldFormat> formatSwitches) {
            return formatSwitches
                .Where(fieldFormat => fieldFormat != WordFieldFormat.Mergeformat && fieldFormat != WordFieldFormat.CharFormat)
                .Cast<WordFieldFormat?>()
                .LastOrDefault();
        }

        private static string CapitalizeFirstWord(string source) {
            if (source.Length == 0) {
                return source;
            }

            char[] chars = source.ToLowerInvariant().ToCharArray();
            for (int i = 0; i < chars.Length; i++) {
                if (char.IsLetter(chars[i])) {
                    chars[i] = char.ToUpperInvariant(chars[i]);
                    break;
                }
            }

            return new string(chars);
        }

        private static string CapitalizeEachWord(string source) {
            if (source.Length == 0) {
                return source;
            }

            char[] chars = source.ToLowerInvariant().ToCharArray();
            bool nextLetterStartsWord = true;
            for (int i = 0; i < chars.Length; i++) {
                if (!char.IsLetter(chars[i])) {
                    nextLetterStartsWord = true;
                    continue;
                }

                if (nextLetterStartsWord) {
                    chars[i] = char.ToUpperInvariant(chars[i]);
                    nextLetterStartsWord = false;
                }
            }

            return new string(chars);
        }

        private static BookmarkStart? FindBookmarkStart(WordDocument document, string bookmarkName) {
            var mainPart = document._wordprocessingDocument.MainDocumentPart;
            if (mainPart == null) {
                return null;
            }

            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                BookmarkStart? bookmarkStart = root.Root
                    .Descendants<BookmarkStart>()
                    .FirstOrDefault(bookmark =>
                        string.Equals(bookmark.Name?.Value, bookmarkName, StringComparison.OrdinalIgnoreCase));
                if (bookmarkStart != null) {
                    return bookmarkStart;
                }
            }

            return null;
        }

        private static string GetBookmarkRangeText(BookmarkStart bookmarkStart) {
            string? bookmarkId = bookmarkStart.Id?.Value;
            if (string.IsNullOrWhiteSpace(bookmarkId)) {
                return string.Empty;
            }

            OpenXmlElement root = bookmarkStart.Ancestors().LastOrDefault() ?? bookmarkStart;
            var text = new List<string>();
            bool inRange = false;
            Paragraph? lastParagraph = null;
            TableCell? lastCell = null;

            foreach (OpenXmlElement element in root.Descendants()) {
                if (ReferenceEquals(element, bookmarkStart)) {
                    inRange = true;
                    continue;
                }

                if (element is BookmarkEnd bookmarkEnd && bookmarkEnd.Id?.Value == bookmarkId) {
                    break;
                }

                if (inRange && element is Text textElement) {
                    Paragraph? paragraph = textElement.Ancestors<Paragraph>().FirstOrDefault();
                    TableCell? cell = textElement.Ancestors<TableCell>().FirstOrDefault();
                    if (text.Count > 0) {
                        if (lastCell != null && cell != null && !ReferenceEquals(lastCell, cell)) {
                            text.Add("\t");
                        } else if (lastParagraph != null && paragraph != null && !ReferenceEquals(lastParagraph, paragraph)) {
                            text.Add("\n");
                        }
                    }

                    text.Add(textElement.Text);
                    lastParagraph = paragraph;
                    lastCell = cell;
                }
            }

            return string.Concat(text);
        }

        private static string GetVisibleParagraphText(Paragraph paragraph) {
            return string.Concat(paragraph.Descendants<Text>().Select(text => text.Text));
        }
    }
}
