using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Biff;

internal static partial class BiffHyperlinkReader {
    /// <summary>Changes only the internal location string, retaining optional fields and native link metadata.</summary>
    internal static byte[] RenameInternalSheet(byte[] payload, IReadOnlyDictionary<string, string> names) {
        if (payload.Length < 32 || BiffRecordReader.ReadUInt32(payload, 24) != StreamVersion) return payload;
        uint flags = BiffRecordReader.ReadUInt32(payload, 28);
        if ((flags & HasLocation) == 0 || (flags & HasMoniker) != 0) return payload;
        int offset = 32;
        if ((flags & HasDisplayName) != 0 && !TrySkipHyperlinkString(payload, ref offset)) throw new InvalidDataException("The retained hyperlink display name is truncated.");
        if ((flags & HasFrameName) != 0 && !TrySkipHyperlinkString(payload, ref offset)) throw new InvalidDataException("The retained hyperlink frame name is truncated.");
        int start = offset;
        if (!TryReadHyperlinkString(payload, ref offset, out string? location)) throw new InvalidDataException("The retained hyperlink location is truncated.");
        string replacement = ExcelDocument.ReplaceSheetNameReferences(location!, names);
        if (replacement == location) return payload;
        byte[] text = Encoding.Unicode.GetBytes(replacement + '\0');
        byte[] result = new byte[checked(start + 4 + text.Length + payload.Length - offset)];
        Buffer.BlockCopy(payload, 0, result, 0, start);
        uint count = checked((uint)(replacement.Length + 1));
        for (int index = 0; index < 4; index++) result[start + index] = (byte)(count >> (8 * index));
        Buffer.BlockCopy(text, 0, result, start + 4, text.Length);
        Buffer.BlockCopy(payload, offset, result, start + 4 + text.Length, payload.Length - offset);
        if (result.Length > 8224) throw new NotSupportedException("The renamed native hyperlink exceeds the BIFF8 record limit.");
        return result;
    }
}
