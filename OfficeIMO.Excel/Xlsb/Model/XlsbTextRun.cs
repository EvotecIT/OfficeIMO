using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Model;

/// <summary>A StrRun (MS-XLSB 2.5.144): the first character of a formatting run and its BrtFont index.</summary>
internal readonly struct XlsbTextRun {
    internal XlsbTextRun(ushort start, ushort font) {
        Start = start;
        Font = font;
    }

    internal ushort Start { get; }

    internal ushort Font { get; }

    /// <summary>
    /// Reads the formatting runs of a RichStr (MS-XLSB 2.5.122) after its flags and text. Returns false when the runs
    /// are malformed or the string carries phonetic data, so the caller keeps the record as preserved native content.
    /// </summary>
    internal static bool TryReadRuns(XlsbBinaryCursor cursor, byte flags, int textLength, out XlsbTextRun[]? runs) {
        runs = null;
        if ((flags & 0x02) != 0) return false;
        if ((flags & 0x01) == 0) return cursor.Remaining == 0;
        if (cursor.Remaining < 4) return false;
        uint count = cursor.ReadUInt32();
        if (count == 0U || count > 0x7FFF || cursor.Remaining != count * 4U) return false;
        var result = new XlsbTextRun[count];
        for (int index = 0; index < result.Length; index++) {
            ushort start = cursor.ReadUInt16();
            ushort font = cursor.ReadUInt16();
            if (start >= textLength || (index > 0 && start <= result[index - 1].Start)) return false;
            result[index] = new XlsbTextRun(start, font);
        }
        runs = result;
        return true;
    }
}
