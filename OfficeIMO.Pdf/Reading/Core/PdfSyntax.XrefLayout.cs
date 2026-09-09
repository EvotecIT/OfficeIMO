namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static IEnumerable<XrefStreamEntry> ReadXrefStreamEntries(PdfDictionary dictionary, byte[] data, Action reportIncomplete) {
        if (!TryReadXrefLayout(dictionary, out int[] widths, out var ranges)) {
            reportIncomplete();
            yield break;
        }
        int entryWidth = widths[0] + widths[1] + widths[2];
        long declaredCount = ranges.Sum(static range => (long)range.Count);
        if (declaredCount * entryWidth > data.LongLength) reportIncomplete();

        int dataOffset = 0;
        foreach (var range in ranges) {
            for (int i = 0; i < range.Count; i++) {
                if (dataOffset > data.Length - entryWidth) yield break;
                long type = widths[0] == 0 ? 1 : ReadBigEndian(data, dataOffset, widths[0]);
                dataOffset += widths[0];
                long field1 = ReadBigEndian(data, dataOffset, widths[1]);
                dataOffset += widths[1];
                long field2 = ReadBigEndian(data, dataOffset, widths[2]);
                dataOffset += widths[2];
                if (type < 0 || type > 2 || field1 < 0 || field2 < 0 ||
                    (type != 0 && (field1 > int.MaxValue || field2 > int.MaxValue))) reportIncomplete();
                yield return new XrefStreamEntry(range.FirstObjectNumber + i, type, field1, field2);
            }
        }
    }

    private static bool TryReadXrefLayout(PdfDictionary dictionary, out int[] widths,
        out List<(int FirstObjectNumber, int Count)> ranges) {
        widths = new int[3];
        ranges = new();
        if (dictionary.Get<PdfArray>("W") is not PdfArray widthsArray || widthsArray.Items.Count != 3) return false;
        for (int index = 0; index < widths.Length; index++)
            if (!TryReadXrefInteger(widthsArray.Items[index], out widths[index]) || widths[index] > sizeof(long)) return false;
        if (widths[0] + widths[1] + widths[2] == 0 ||
            !TryReadXrefInteger(dictionary.Get<PdfNumber>("Size"), out int size) || size == 0) return false;

        if (dictionary.Items.TryGetValue("Index", out PdfObject? indexValue)) {
            if (indexValue is not PdfArray indexArray || indexArray.Items.Count == 0 || indexArray.Items.Count % 2 != 0) return false;
            for (int index = 0; index < indexArray.Items.Count; index += 2) {
                if (!TryReadXrefInteger(indexArray.Items[index], out int first) ||
                    !TryReadXrefInteger(indexArray.Items[index + 1], out int count) || (long)first + count > size) return false;
                if (count > 0) ranges.Add((first, count));
            }
            long previousEnd = -1;
            foreach (var range in ranges.OrderBy(static range => range.FirstObjectNumber)) {
                if (range.FirstObjectNumber < previousEnd) return false;
                previousEnd = (long)range.FirstObjectNumber + range.Count;
            }
        } else ranges.Add((0, size));
        return ranges.Count > 0;
    }

    private static bool TryReadXrefInteger(PdfObject? value, out int integer) {
        integer = 0;
        if (value is not PdfNumber number || double.IsNaN(number.Value) ||
            number.Value < 0 || number.Value > int.MaxValue || number.Value != Math.Truncate(number.Value)) return false;
        integer = (int)number.Value;
        return true;
    }

    private static long ReadBigEndian(byte[] data, int offset, int length) {
        long value = 0;
        for (int index = 0; index < length; index++) value = (value << 8) | data[offset + index];
        return value;
    }
}
