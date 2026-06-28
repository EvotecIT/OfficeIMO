using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private static byte[] BuildSelectionPayload(LegacyXlsSelection selection) {
            using var stream = new MemoryStream();
            stream.WriteByte(selection.Pane);
            WriteUInt16(stream, checked((ushort)(selection.ActiveRow - 1)));
            WriteUInt16(stream, checked((ushort)(selection.ActiveColumn - 1)));
            WriteUInt16(stream, selection.ActiveRangeIndex);
            WriteUInt16(stream, checked((ushort)selection.SelectedRanges.Count));
            foreach (LegacyXlsSelectedRange range in selection.SelectedRanges) {
                WriteUInt16(stream, checked((ushort)(range.StartRow - 1)));
                WriteUInt16(stream, checked((ushort)(range.EndRow - 1)));
                stream.WriteByte(checked((byte)(range.StartColumn - 1)));
                stream.WriteByte(checked((byte)(range.EndColumn - 1)));
            }

            return stream.ToArray();
        }

        private static IReadOnlyList<LegacyXlsSelection> ExtractSelections(SheetView? sheetView, int frozenRowCount, int frozenColumnCount, LegacyXlsSplitPaneView? splitPane) {
            var selections = new List<LegacyXlsSelection>();
            if (sheetView != null) {
                foreach (Selection selection in sheetView.Elements<Selection>()) {
                    if (TryCreateSelection(selection, out LegacyXlsSelection? legacySelection)) {
                        selections.Add(legacySelection!);
                    }
                }
            }

            if (selections.Count == 0 && (frozenRowCount > 0 || frozenColumnCount > 0)) {
                int activeRow = Math.Max(1, frozenRowCount + 1);
                int activeColumn = Math.Max(1, frozenColumnCount + 1);
                selections.Add(new LegacyXlsSelection(
                    ResolveActivePane(frozenRowCount, frozenColumnCount),
                    activeRow,
                    activeColumn,
                    0,
                    new[] { new LegacyXlsSelectedRange(activeRow, activeColumn, activeRow, activeColumn) }));
            } else if (selections.Count == 0 && splitPane.HasValue) {
                int activeRow = Math.Max(1, splitPane.Value.TopRow + 1);
                int activeColumn = Math.Max(1, splitPane.Value.LeftColumn + 1);
                selections.Add(new LegacyXlsSelection(
                    splitPane.Value.ActivePane,
                    activeRow,
                    activeColumn,
                    0,
                    new[] { new LegacyXlsSelectedRange(activeRow, activeColumn, activeRow, activeColumn) }));
            }

            return selections;
        }

        private static bool TryCreateSelection(Selection selection, out LegacyXlsSelection? legacySelection) {
            legacySelection = null;
            string activeCell = selection.ActiveCell?.Value ?? "A1";
            if (!A1.TryParseCellReferenceFast(activeCell, out int activeRow, out int activeColumn)
                || !IsWithinBiff8Limits(activeRow, activeColumn)) {
                return false;
            }

            IReadOnlyList<LegacyXlsSelectedRange> selectedRanges = ParseSelectedRanges(selection.SequenceOfReferences?.InnerText, activeRow, activeColumn);
            if (selectedRanges.Count == 0) {
                return false;
            }

            legacySelection = new LegacyXlsSelection(
                ToLegacyPane(selection.Pane?.Value),
                activeRow,
                activeColumn,
                0,
                selectedRanges);
            return true;
        }

        private static IReadOnlyList<LegacyXlsSelectedRange> ParseSelectedRanges(string? sequenceOfReferences, int fallbackRow, int fallbackColumn) {
            if (string.IsNullOrWhiteSpace(sequenceOfReferences)) {
                return new[] { new LegacyXlsSelectedRange(fallbackRow, fallbackColumn, fallbackRow, fallbackColumn) };
            }

            var ranges = new List<LegacyXlsSelectedRange>();
            foreach (string token in sequenceOfReferences!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (TryParseSelectedRange(token, out LegacyXlsSelectedRange? selectedRange)) {
                    ranges.Add(selectedRange!);
                }
            }

            return ranges.Count == 0
                ? new[] { new LegacyXlsSelectedRange(fallbackRow, fallbackColumn, fallbackRow, fallbackColumn) }
                : ranges;
        }

        private static bool TryParseSelectedRange(string reference, out LegacyXlsSelectedRange? selectedRange) {
            selectedRange = null;
            string normalized = reference.Replace("$", string.Empty);
            if (A1.TryParseCellReferenceFast(normalized, out int row, out int column)) {
                if (!IsWithinBiff8Limits(row, column)) {
                    return false;
                }

                selectedRange = new LegacyXlsSelectedRange(row, column, row, column);
                return true;
            }

            if (!A1.TryParseRange(normalized, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
                || !IsWithinBiff8Limits(firstRow, firstColumn)
                || !IsWithinBiff8Limits(lastRow, lastColumn)) {
                return false;
            }

            selectedRange = new LegacyXlsSelectedRange(firstRow, firstColumn, lastRow, lastColumn);
            return true;
        }

        private static bool IsWithinBiff8Limits(int row, int column) {
            return row >= 1 && row <= 65536 && column >= 1 && column <= 256;
        }

        private static byte ToLegacyPane(PaneValues? pane) {
            if (!pane.HasValue) return 3;
            if (pane.Value == PaneValues.BottomRight) return 0;
            if (pane.Value == PaneValues.TopRight) return 1;
            if (pane.Value == PaneValues.BottomLeft) return 2;
            return 3;
        }
    }
}
