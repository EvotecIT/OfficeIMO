using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.Legacy;

internal sealed partial class EarlyBiffWorkbook {
    /// <summary>Collects BIFF4 group declarations and their cached formula members during the existing record scan.</summary>
    private void ObserveArrayFormula(Record record, ref ArrayFormulaState? state) {
        if (record.Type == 0x0221) {
            state ??= new ArrayFormulaState();
            if (state.Records.Count >= _limits.MaxItems) throw new InvalidDataException("Early BIFF exceeds the configured array formula limit.");
            state.Records.Add(record);
        } else if (record.Type == 0x0406 && TryGetArrayAnchor(record, out int anchor)) {
            state ??= new ArrayFormulaState();
            if (!state.Members.TryGetValue(anchor, out List<int>? members)) {
                members = new List<int>();
                state.Members.Add(anchor, members);
            }
            if (++state.MemberCount > _limits.MaxItems) throw new InvalidDataException("Early BIFF exceeds the configured array formula cell limit.");
            members.Add(ReadUInt16(record, 0) * MaximumColumns + ReadUInt16(record, 2));
        }
    }

    /// <summary>Accepts complete groups only, so projecting an array cannot erase unrelated formulas or manufacture missing cells.</summary>
    private void PrepareArrayFormulas(ArrayFormulaState? state, LegacySpreadsheetSheet sheet, Encoding encoding) {
        if (state == null) return;
        var anchors = new HashSet<int>();
        foreach (Record record in state.Records) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (record.Length < 10) throw new InvalidDataException("Truncated early BIFF array formula.");
            int firstRow = ReadUInt16(record, 0);
            int lastRow = ReadUInt16(record, 2);
            int firstColumn = _data[record.Offset + 4];
            int lastColumn = _data[record.Offset + 5];
            if (lastRow < firstRow || lastRow >= MaximumRows || lastColumn < firstColumn) {
                throw new InvalidDataException("Invalid early BIFF array formula range.");
            }
            int anchor = firstRow * MaximumColumns + firstColumn;
            if (!anchors.Add(anchor)) throw new InvalidDataException("Duplicate early BIFF array formula anchor.");
            int count = ReadUInt16(record, 8);
            if (count > record.Length - 10) throw new InvalidDataException("Truncated early BIFF array formula tokens.");
            int cells = (lastRow - firstRow + 1) * (lastColumn - firstColumn + 1);
            if (!state.Members.TryGetValue(anchor, out List<int>? members) || members.Count != cells) continue;
            bool complete = true;
            foreach (int member in members) {
                _cancellationToken.ThrowIfCancellationRequested();
                int row = member / MaximumColumns;
                int column = member % MaximumColumns;
                if (row < firstRow || row > lastRow || column < firstColumn || column > lastColumn) {
                    complete = false;
                    break;
                }
            }
            // Ancillary array-constant data remains outside this qualified profile.
            if (!complete || count != record.Length - 10
                || !EarlyBiffFormulaReader.TryRead(_data, record.Offset + 10, count, 0x0409,
                    encoding, firstRow, firstColumn, out string? formula)) continue;
            var array = new LegacyXlsArrayFormulaRecord(firstRow + 1, firstColumn + 1, lastRow + 1, lastColumn + 1,
                ReadUInt16(record, 6), count, 0, members.Count, true, record.Offset - HeaderSize, record.Type, record.Length);
            state.Projected.Add(anchor, (array, formula!));
            state.ProjectedRecords.Add(record.Offset);
            sheet.ArrayFormulas.Add((array, formula!));
        }
    }

    private bool TryGetArrayAnchor(Record record, out int anchor) {
        anchor = 0;
        if (record.Length < 23 || ReadUInt16(record, 16) != 5 || _data[record.Offset + 18] != 0x01) return false;
        int row = ReadUInt16(record, 19);
        int column = ReadUInt16(record, 21);
        if (row >= MaximumRows || column >= MaximumColumns) return false;
        anchor = row * MaximumColumns + column;
        return true;
    }

    private static bool TryReadArrayFormula(ArrayFormulaState? state, int anchor, int row, int column, out string? formula) {
        formula = null;
        if (state == null || !state.Projected.TryGetValue(anchor, out var array)) return false;
        LegacyXlsArrayFormulaRecord range = array.Record;
        if (row + 1 < range.FirstRow || row + 1 > range.LastRow || column + 1 < range.FirstColumn || column + 1 > range.LastColumn) return false;
        if (row + 1 == range.FirstRow && column + 1 == range.FirstColumn) formula = array.Formula;
        return true;
    }

    private void ObserveArrayOrUnprojected(Record record, ArrayFormulaState? state) {
        if (state == null || !state.ProjectedRecords.Contains(record.Offset)) ObserveUnprojected(record);
    }

    private sealed class ArrayFormulaState {
        internal List<Record> Records { get; } = new();
        internal Dictionary<int, List<int>> Members { get; } = new();
        internal Dictionary<int, (LegacyXlsArrayFormulaRecord Record, string Formula)> Projected { get; } = new();
        internal HashSet<int> ProjectedRecords { get; } = new();
        internal int MemberCount { get; set; }
    }
}
