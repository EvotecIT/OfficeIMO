using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write;

internal static partial class XlsbWorksheetPartWriter {
    private static IReadOnlyList<XlsbWriteCell> PrepareRetainedFormulaGroups(IReadOnlyList<XlsbRecord> records, IReadOnlyList<XlsbWriteCell> cells) {
        var result = cells.ToDictionary(cell => (cell.Row, cell.Column));
        var groups = XlsbFormulaGroup.ReadGroups(records);
        var index = new ExcelCellRangeIndex<XlsbWriteCell>(cells.Where(cell => XlsbFormulaGroup.IsGroupReference(cell.FormulaPayload) || cell.FormulaGroup != null)
            .Select(cell => (cell.Row, cell.Column, cell.Row, cell.Column, cell)));
        foreach (var group in groups) {
            var covered = index.Intersect(group.Range.FirstRow, group.Range.FirstColumn, group.Range.LastRow, group.Range.LastColumn).ToArray();
            var members = covered.Where(cell => XlsbFormulaGroup.IsGroupReference(cell.FormulaPayload)).ToArray();
            if (group.Type == 427 && members.Length == group.SourceMemberCount
                && result.TryGetValue((group.AnchorRow, group.AnchorColumn), out var sharedOwner)
                && XlsbFormulaGroup.IsGroupReference(sharedOwner.FormulaPayload)) {
                sharedOwner.FormulaGroup = new XlsbGeneratedRecord(427, group.Data);
            } else if (group.Type == 427) {
                foreach (var cell in members) {
                    byte[] payload = group.Materialize(cell.Row, cell.Column);
                    if (cell.FormulaPayload != null) { payload[0] = cell.FormulaPayload[0]; payload[1] = cell.FormulaPayload[1]; }
                    result[(cell.Row, cell.Column)] = WithFormula(cell, payload);
                }
            } else if (members.Length != 0 && !covered.Any(cell => cell.FormulaGroup != null)) {
                long expected = (long)(group.Range.LastRow - group.Range.FirstRow + 1) * (group.Range.LastColumn - group.Range.FirstColumn + 1);
                if (members.Length != expected) throw new NotSupportedException("Clear or replace the complete native XLSB array range before editing its members.");
                if (!result.TryGetValue((group.AnchorRow, group.AnchorColumn), out var owner)) throw new InvalidDataException("The native XLSB array owner is missing.");
                owner.FormulaGroup = new XlsbGeneratedRecord(426, group.Data);
            }
        }
        return result.Values.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToArray();
    }

    internal static XlsbWriteCell WithFormula(XlsbWriteCell cell, byte[] formula) {
        XlsbWriteCellKind kind = cell.Kind switch {
            XlsbWriteCellKind.Number => XlsbWriteCellKind.FormulaNumber,
            XlsbWriteCellKind.Boolean => XlsbWriteCellKind.FormulaBoolean,
            XlsbWriteCellKind.Error => XlsbWriteCellKind.FormulaError,
            XlsbWriteCellKind.Text or XlsbWriteCellKind.Blank => XlsbWriteCellKind.FormulaText,
            _ => cell.Kind
        };
        object value = cell.Value ?? string.Empty;
        if (kind == XlsbWriteCellKind.FormulaError && value is string error) {
            if (!ExcelErrorCode.TryGetCode(error, out byte code)) throw new InvalidDataException("The native XLSB formula cache has an invalid error value.");
            value = code;
        }
        return new XlsbWriteCell(cell.Row, cell.Column, cell.StyleIndex, kind, value, formula) { FormulaGroup = cell.FormulaGroup };
    }
}
