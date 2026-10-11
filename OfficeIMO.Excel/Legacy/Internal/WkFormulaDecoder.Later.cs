using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Excel.Legacy;

internal sealed class LaterFormulaContext {
    internal bool IsLotus { get; set; }
    internal bool ExtendedNumbers { get; set; }
    internal int Sheet { get; set; }
    internal Func<int, string>? SheetName { get; set; }
    internal IReadOnlyList<(string Text, bool Range)>? References { get; set; }
    private int _reference;
    internal bool ReferencesConsumed => References == null || _reference == References.Count;
    internal static string Range(string first, string last) {
        int firstPrefix = first.LastIndexOf('!') + 1, lastPrefix = last.LastIndexOf('!') + 1;
        if (!string.Equals(first.Substring(0, firstPrefix), last.Substring(0, lastPrefix), StringComparison.Ordinal))
            throw new InvalidDataException("A range across distinct sheets is outside the qualified formula profile.");
        return first + ":" + last.Substring(lastPrefix);
    }
    internal string NextReference(bool range) {
        if (References == null || _reference >= References.Count || References[_reference].Range != range)
            throw new InvalidDataException("Formula reference table does not match the token stream.");
        return References[_reference++].Text;
    }
}

internal static partial class WkFormulaDecoder {
    private static int ReadLaterFlags(byte[] data, ref int cursor, int end) { Require(cursor, 1, end); return data[cursor++]; }
    private static double ReadLaterNumber(byte[] data, ref int cursor, int end, LaterFormulaContext? context, bool compact) {
        int size = compact ? (context?.IsLotus == true && !context.ExtendedNumbers ? 4 : 2) :
            context?.ExtendedNumbers == true ? 10 : 8;
        Require(cursor, size, end);
        double value = context?.IsLotus == true ? LaterSpreadsheetNumbers.Read(data, cursor,
            compact ? (size == 2 ? LaterNumberKind.Compact16 : LaterNumberKind.Compact32) :
            size == 10 ? LaterNumberKind.Extended80 : LaterNumberKind.Double64) : compact ?
            (short)(data[cursor] | data[cursor + 1] << 8) : BitConverter.IsLittleEndian ? BitConverter.ToDouble(data, cursor) : ReadBigEndianDouble(data, cursor);
        cursor += size;
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new InvalidDataException("Formula numeric token is not finite.");
        return value;
    }
    private static double ReadBigEndianDouble(byte[] data, int offset) { var copy = new byte[8]; Array.Copy(data, offset, copy, 0, 8); Array.Reverse(copy); return BitConverter.ToDouble(copy, 0); }
    private static string ReadLaterReference(byte[] data, ref int cursor, int end, int row, int column, LaterFormulaContext? context, int flags = -1) {
        if (context?.References != null) return context.NextReference(range: false);
        if (context?.IsLotus != true) return ReadReference(data, ref cursor, end, row, column);
        if (flags < 0) flags = ReadLaterFlags(data, ref cursor, end);
        if ((flags & ~3) != 0) throw new InvalidDataException("Unsupported Lotus formula reference flags.");
        Require(cursor, 4, end);
        int targetRow = data[cursor] | data[cursor + 1] << 8, sheet = data[cursor + 2], targetColumn = data[cursor + 3]; cursor += 4;
        string prefix = sheet == context.Sheet ? string.Empty : "'" + context.SheetName!(sheet).Replace("'", "''") + "'!";
        return prefix + ((flags & 1) == 0 ? "$" : "") + ColumnName(targetColumn + 1) +
            ((flags & 2) == 0 ? "$" : "") + (targetRow + 1).ToString(CultureInfo.InvariantCulture);
    }
}
