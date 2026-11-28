using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using System.Globalization;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private static bool TryEvaluateRef(
            WordDocument document,
            MutableFieldCandidate candidate,
            WordFieldInventory.ParsedFieldInstruction parsed,
            FieldEvaluationState state,
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

            OfficeXmlNode bookmarkStart = FindBookmarkStart(document, bookmarkName!);
            if (bookmarkStart.IsNone) {
                message = $"Bookmark {bookmarkName} was not found.";
                return false;
            }

            if (listSwitch.HasValue) {
                if (!TryResolveReferenceListNumber(document, candidate, bookmarkStart, state, listSwitch.Value, parsed.Switches, out string? listNumber, out WordFieldUpdateStatus listStatus, out string listMessage)) {
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
                OfficeXmlNode paragraph = bookmarkStart.AncestorOrSelf(OfficeXmlNamespaces.WordprocessingML, "p");
                bookmarkText = string.Concat(WordFieldInventory.EnumerateFieldContent(paragraph)
                    .Where(node => node.Is(OfficeXmlNamespaces.WordprocessingML, "t")).Select(node => node.InnerText));
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
            message = parsed.FormatSwitches.Any(fieldFormat => fieldFormat != WordFieldFormat.MergeFormat && fieldFormat != WordFieldFormat.CharFormat)
                ? $"Updated from bookmark {bookmarkName} with deterministic text formatting."
                : $"Updated from bookmark {bookmarkName}.";
            return true;
        }

        private static bool TryEvaluatePageRef(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            WordLayoutPagination? layout,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message,
            out bool fromLayout) {
            fromLayout = false;
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

            OfficeXmlNode bookmarkStart = FindBookmarkStart(document, bookmarkName!);
            if (bookmarkStart.IsNone) {
                message = $"Bookmark {bookmarkName} was not found.";
                return false;
            }

            if (bookmarkStart.AncestorOrSelf(OfficeXmlNamespaces.WordprocessingML, "body").IsNone) {
                message = "PAGEREF fields outside the document body need Word layout context and were left unchanged.";
                return false;
            }

            int? page = null;
            string? layoutLabel = null;
            if (layout != null) {
                if (!layout.TryGetBookmarkPage(bookmarkStart.Attribute(OfficeXmlNamespaces.WordprocessingML, "name") ?? bookmarkName!, out int laidOutPage, out layoutLabel)) {
                    // A page-break estimate would contradict the laid out pages, so keep the existing result.
                    message = $"Bookmark {bookmarkName} could not be placed on a page by OfficeIMO layout; the PAGEREF result was left unchanged.";
                    return false;
                }

                page = laidOutPage;
            }

            fromLayout = page != null;
            page ??= EstimatePageForBodyField(document, bookmarkStart);
            if (page == null) {
                message = $"Bookmark {bookmarkName} could not be matched to a body paragraph.";
                return false;
            }

            if (!TryFormatPageReferenceValue(page.Value, parsed.FormatSwitches, parsed.NumericPictureSwitch, out string pageText, out string? unsupportedFormat, out bool usedNumericPicture)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = $"PAGEREF field format switch {unsupportedFormat} is not supported for deterministic page-reference refresh.";
                return false;
            }

            if (fromLayout && string.IsNullOrWhiteSpace(parsed.NumericPictureSwitch)
                && GetLastMeaningfulFormat(parsed.FormatSwitches) == null) {
                if (layoutLabel == null) {
                    status = WordFieldUpdateStatus.Unsupported;
                    message = "The target section page-number format is not supported and the PAGEREF result was left unchanged.";
                    return false;
                }
                pageText = layoutLabel;
            }

            value = pageText;
            status = WordFieldUpdateStatus.Updated;
            string source = fromLayout ? "OfficeIMO page layout" : "OfficeIMO page-break order";
            message = usedNumericPicture
                ? $"Updated from {source} for bookmark {bookmarkName} with numeric picture formatting."
                : parsed.FormatSwitches.Any(fieldFormat => fieldFormat != WordFieldFormat.MergeFormat && fieldFormat != WordFieldFormat.CharFormat)
                ? $"Updated from {source} for bookmark {bookmarkName} with deterministic number formatting."
                : $"Updated from {source} for bookmark {bookmarkName}.";
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

        internal static bool TryApplyReferenceTextFormat(
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
                case WordFieldFormat.UpperRoman:
                case WordFieldFormat.LowerRoman:
                case WordFieldFormat.Ordinal:
                case WordFieldFormat.LowerLetter:
                case WordFieldFormat.UpperLetter:
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
                .Where(fieldFormat => fieldFormat != WordFieldFormat.MergeFormat && fieldFormat != WordFieldFormat.CharFormat)
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

        private static OfficeXmlNode FindBookmarkStart(WordDocument document, string bookmarkName) {
            const int W = OfficeXmlNamespaces.WordprocessingML;
            foreach (var story in WordFieldInventory.EnumerateStories(document)) {
                foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(story.Root)) {
                    if (node.Is(W, "bookmarkStart") &&
                        string.Equals(node.Attribute(W, "name"), bookmarkName, StringComparison.OrdinalIgnoreCase))
                        return node;
                }
            }
            return default;
        }

        private static string GetBookmarkRangeText(OfficeXmlNode bookmarkStart) {
            const int W = OfficeXmlNamespaces.WordprocessingML;
            string? bookmarkId = bookmarkStart.Attribute(W, "id");
            if (string.IsNullOrWhiteSpace(bookmarkId)) return string.Empty;

            OfficeXmlNode root = bookmarkStart;
            while (root.Parent.Exists) root = root.Parent;
            var text = new List<string>();
            bool inRange = false;
            OfficeXmlNode lastParagraph = default;
            OfficeXmlNode lastCell = default;

            foreach (OfficeXmlNode element in WordFieldInventory.EnumerateFieldContent(root)) {
                if (element == bookmarkStart) {
                    inRange = true;
                    continue;
                }

                if (element.Is(W, "bookmarkEnd") && element.Attribute(W, "id") == bookmarkId) break;
                if (!inRange || !element.Is(W, "t")) continue;

                OfficeXmlNode paragraph = element.AncestorOrSelf(W, "p");
                OfficeXmlNode cell = element.AncestorOrSelf(W, "tc");
                if (text.Count > 0) {
                    if (lastCell.Exists && cell.Exists && lastCell != cell) {
                        text.Add("\t");
                    } else if (lastParagraph.Exists && paragraph.Exists && lastParagraph != paragraph) {
                        text.Add("\n");
                    }
                }

                text.Add(element.InnerText);
                lastParagraph = paragraph;
                lastCell = cell;
            }

            return string.Concat(text);
        }
    }
}
