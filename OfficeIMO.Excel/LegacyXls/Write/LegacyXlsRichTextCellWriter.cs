using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsRichTextCellWriter {
        internal static bool SupportsWorksheetCellTextRuns(
            WorkbookPart workbookPart,
            Worksheet worksheet,
            LegacyXlsFontTable fontTable,
            out string? reason) {
            reason = null;
            SharedStringTable? sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
            foreach (Cell cell in worksheet.Descendants<Cell>()) {
                if (!TryGetCellRichTextRuns(cell, sharedStringTable, fontTable, out _, out _, out reason)) {
                    return false;
                }
            }

            return true;
        }

        internal static bool TryGetCellRichTextRuns(
            Cell cell,
            SharedStringTable? sharedStringTable,
            LegacyXlsFontTable fontTable,
            out string? text,
            out IReadOnlyList<LegacyXlsTextFormattingRun> formattingRuns,
            out string? reason) {
            text = null;
            formattingRuns = Array.Empty<LegacyXlsTextFormattingRun>();
            reason = null;

            if (cell.InlineString != null) {
                return TryGetInlineStringRichTextRuns(cell.InlineString, fontTable, cell.CellReference?.Value, out text, out formattingRuns, out reason);
            }

            if (cell.DataType?.Value == CellValues.SharedString
                && TryGetSharedStringItem(sharedStringTable, cell.CellValue?.InnerText, out SharedStringItem? sharedStringItem)) {
                return TryGetSharedStringRichTextRuns(sharedStringItem!, fontTable, cell.CellReference?.Value, out text, out formattingRuns, out reason);
            }

            return true;
        }

        private static bool TryGetInlineStringRichTextRuns(
            InlineString inlineString,
            LegacyXlsFontTable fontTable,
            string? cellReference,
            out string? text,
            out IReadOnlyList<LegacyXlsTextFormattingRun> formattingRuns,
            out string? reason) {
            if (inlineString.Elements<PhoneticRun>().Any() || inlineString.Elements<PhoneticProperties>().Any()) {
                text = null;
                formattingRuns = Array.Empty<LegacyXlsTextFormattingRun>();
                reason = BuildCellTextRunReason(cellReference, "phonetic cell text");
                return false;
            }

            return TryCollectRuns(inlineString.Elements<Run>(), fontTable, out text, out formattingRuns, out reason);
        }

        private static bool TryGetSharedStringRichTextRuns(
            SharedStringItem item,
            LegacyXlsFontTable fontTable,
            string? cellReference,
            out string? text,
            out IReadOnlyList<LegacyXlsTextFormattingRun> formattingRuns,
            out string? reason) {
            if (item.Elements<PhoneticRun>().Any() || item.Elements<PhoneticProperties>().Any()) {
                text = null;
                formattingRuns = Array.Empty<LegacyXlsTextFormattingRun>();
                reason = BuildCellTextRunReason(cellReference, "phonetic cell text");
                return false;
            }

            return TryCollectRuns(item.Elements<Run>(), fontTable, out text, out formattingRuns, out reason);
        }

        private static bool TryCollectRuns(
            IEnumerable<Run> runs,
            LegacyXlsFontTable fontTable,
            out string? text,
            out IReadOnlyList<LegacyXlsTextFormattingRun> formattingRuns,
            out string? reason) {
            text = null;
            formattingRuns = Array.Empty<LegacyXlsTextFormattingRun>();
            reason = null;

            var builder = new StringBuilder();
            var collectedRuns = new List<LegacyXlsTextFormattingRun>();
            foreach (Run run in runs) {
                if (!SupportsRunMetadata(run, out reason)) {
                    return false;
                }

                string runText = run.Text?.Text ?? string.Empty;
                if (runText.Length == 0) {
                    continue;
                }

                if (builder.Length + runText.Length > 32767) {
                    reason = "rich-text cell lengths outside BIFF8 limits";
                    return false;
                }

                if (!fontTable.TryGetFontIndex(run.RunProperties, out ushort fontIndex, out reason)) {
                    return false;
                }

                ushort startCharacter = checked((ushort)builder.Length);
                if (collectedRuns.Count == 0 || collectedRuns[collectedRuns.Count - 1].FontIndex != fontIndex) {
                    collectedRuns.Add(new LegacyXlsTextFormattingRun(startCharacter, fontIndex));
                }

                builder.Append(runText);
            }

            if (collectedRuns.Count == 0) {
                return true;
            }

            text = builder.ToString();
            formattingRuns = collectedRuns;
            return true;
        }

        private static bool SupportsRunMetadata(Run run, out string? reason) {
            reason = null;
            if (run.GetAttributes().Any()) {
                reason = "rich-text cell run metadata";
                return false;
            }

            if (run.ChildElements.Any(child => child is not RunProperties && child is not Text)) {
                reason = "rich-text cell run metadata";
                return false;
            }

            if (run.Elements<Text>().Take(2).Count() > 1) {
                reason = "rich-text cell run metadata";
                return false;
            }

            return true;
        }

        private static bool TryGetSharedStringItem(SharedStringTable? sharedStringTable, string? indexText, out SharedStringItem? item) {
            item = null;
            if (sharedStringTable == null || string.IsNullOrWhiteSpace(indexText)) {
                return false;
            }

            if (!int.TryParse(indexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) || index < 0) {
                return false;
            }

            item = sharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(index);
            return item != null;
        }

        private static string BuildCellTextRunReason(string? cellReference, string feature) {
            string reference = string.IsNullOrWhiteSpace(cellReference) ? "a worksheet cell" : cellReference!;
            return $"{feature} at {reference}";
        }
    }
}
