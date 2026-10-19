using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Biff12;

/// <summary>Owns native shared/array definitions and their position-dependent token materialization.</summary>
internal sealed class XlsbFormulaGroup {
    internal XlsbFormulaGroup(int type, byte[] data, int row, int column) {
        Type = type; Data = data; AnchorRow = row; AnchorColumn = column;
        int prefix = type == 426 ? 17 : 16;
        if (data.Length < prefix + 8) throw new InvalidDataException("The XLSB formula group is truncated.");
        Range = new XlsbCellRange(checked((int)Read(data, 0) + 1), checked((int)Read(data, 4) + 1),
            checked((int)Read(data, 8) + 1), checked((int)Read(data, 12) + 1));
        if (Range.FirstRow < 1 || Range.LastRow < Range.FirstRow || Range.LastRow > 1_048_576
            || Range.FirstColumn < 1 || Range.LastColumn < Range.FirstColumn || Range.LastColumn > 16_384 || !Contains(row, column))
            throw new InvalidDataException("The XLSB formula group range is invalid.");
        int count = checked((int)Read(data, prefix));
        if (count == 0 || count > data.Length - prefix - 8) throw new InvalidDataException("The XLSB formula-group token range is invalid.");
        int extra = checked((int)Read(data, prefix + 4 + count));
        if (extra != data.Length - prefix - 8 - count) throw new InvalidDataException("The XLSB formula-group extra-data range is invalid.");
        ParsedFormula = data.Skip(prefix).ToArray();
    }
    internal int SourceMemberCount { get; private set; }
    internal int Type { get; }
    internal byte[] Data { get; }
    internal byte[] ParsedFormula { get; }
    internal XlsbCellRange Range { get; }
    internal int AnchorRow { get; }
    internal int AnchorColumn { get; }
    internal bool Contains(int row, int column) => row >= Range.FirstRow && row <= Range.LastRow && column >= Range.FirstColumn && column <= Range.LastColumn;

    internal byte[] Materialize(int row, int column) {
        byte[] result = new byte[ParsedFormula.Length + 2]; Buffer.BlockCopy(ParsedFormula, 0, result, 2, ParsedFormula.Length);
        if (Type != 427) return result;
        int count = checked((int)Read(result, 2));
        XlsbFormulaTokenScanner.Visit(result, 6, count, (offset, token, operand) => {
            int kind = token & 0x1f;
            if (token < 0x20 || kind is not 0x0c and not 0x0d) return;
            result[offset] = (byte)(token - 8);
            if (kind == 0x0c) Resolve(operand, operand + 4);
            else { Resolve(operand, operand + 8); Resolve(operand + 4, operand + 10); }
        });
        return result;
        void Resolve(int rowOffset, int columnOffset) {
            uint nativeRow = Read(result, rowOffset); int bits = result[columnOffset] | result[columnOffset + 1] << 8;
            int nativeColumn = bits & 0x3fff;
            if ((bits & 0x8000) != 0) nativeRow = unchecked((uint)(row - 1 + (int)nativeRow)) & 0xfffff;
            if ((bits & 0x4000) != 0) nativeColumn = (column - 1 + (nativeColumn >= 0x2000 ? nativeColumn - 0x4000 : nativeColumn)) & 0x3fff;
            Put(result, rowOffset, nativeRow); int value = (bits & 0xc000) | nativeColumn;
            result[columnOffset] = (byte)value; result[columnOffset + 1] = (byte)(value >> 8);
        }
    }
    internal static IReadOnlyList<XlsbFormulaGroup> ReadGroups(IReadOnlyList<XlsbRecord> records) {
        var formulas = new List<(int Row, int Column)>();
        var groups = new List<XlsbFormulaGroup>(); int row = 0, column = 0, futureDepth = 0; bool inData = false, previousFormula = false;
        foreach (var record in records) {
            if (record.Type == 35) { futureDepth++; continue; }
            if (record.Type == 36) { futureDepth--; continue; }
            if (futureDepth > 0) continue;
            if (record.Type == 145) { inData = true; continue; }
            if (record.Type == 146) { inData = false; continue; }
            if (!inData) continue;
            if (record.Type == 0) { row = checked((int)Read(record.Data, 0) + 1); previousFormula = false; }
            else if (record.Type is >= 1 and <= 11 or 62) {
                column = checked((int)Read(record.Data, 0) + 1); previousFormula = record.Type is >= 8 and <= 11;
                if (previousFormula) {
                    int offset = record.Type switch { 8 => checked(12 + (int)Read(record.Data, 8) * 2), 9 => 16, _ => 9 };
                    if (offset <= record.Data.Length - 15 && Read(record.Data, offset + 2) == 5 && record.Data[offset + 6] == 1) formulas.Add((row, column));
                }
            } else if (record.Type is 426 or 427) {
                if (!previousFormula) throw new InvalidDataException("The XLSB formula group has no preceding formula owner.");
                var group = new XlsbFormulaGroup(record.Type, record.Data, row, column);
                if (groups.Any(old => old.Range.FirstRow <= group.Range.LastRow && old.Range.LastRow >= group.Range.FirstRow
                    && old.Range.FirstColumn <= group.Range.LastColumn && old.Range.LastColumn >= group.Range.FirstColumn))
                    throw new InvalidDataException("The XLSB formula group ranges overlap.");
                groups.Add(group); previousFormula = false;
            } else previousFormula = false;
        }
        foreach (var group in groups) group.SourceMemberCount = formulas.Count(cell => group.Contains(cell.Row, cell.Column));
        return groups;
    }
    internal static bool IsGroupReference(byte[]? payload) => payload != null && payload.Length >= 15 && Read(payload, 2) == 5 && payload[6] == 1;
    internal static uint Read(byte[] data, int offset) => XlsbFormulaTokenScanner.ReadUInt32(data, offset);
    internal static void Put(byte[] data, int offset, uint value) => XlsbFormulaTokenScanner.WriteUInt32(data, offset, value);
}
