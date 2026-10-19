using OfficeIMO.Excel.LegacyXls.Biff;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal sealed partial class LegacyXlsExternSheetTable {
    internal static LegacyXlsExternSheetTable CreatePreserving(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets,
        IReadOnlyList<BiffRecord> source, IReadOnlyDictionary<string, int> nativeSheets) {
        var generated = Create(document, sheets);
        var old = source.Where(record => record.Type == 0x0017).ToArray();
        if (old.Length > 1) throw new NotSupportedException("Multiple native ExternSheet records require consolidation.");
        byte[] original = old.Length == 0 ? new byte[2] : old[0].Payload;
        int count = BiffRecordReader.ReadUInt16(original, 0);
        if (original.Length != 2 + count * 6) throw new NotSupportedException("Invalid or continued native ExternSheet table.");
        var books = source.Where(record => record.Type == 0x01ae).ToArray();
        var selfBooks = new HashSet<int>(books.Select((record, index) => (record, index))
            .Where(item => item.record.Payload.Length == 4 && item.record.Payload[2] == 1 && item.record.Payload[3] == 4).Select(item => item.index));
        var entries = new List<ExternSheetEntry>();
        for (int index = 0; index < count; index++) entries.Add(ReadEntry(original, index));
        int generatedCount = BiffRecordReader.ReadUInt16(generated.Payload, 0);
        var nativeEntries = new List<ExternSheetEntry>(); var indexMap = new Dictionary<ushort, ushort>(); bool needsSelf = false;
        for (int index = 0; index < generatedCount; index++) {
            var entry = ReadEntry(generated.Payload, index);
            if (entry.SupBookIndex == 0) {
                entry = new ExternSheetEntry(0, checked((ushort)nativeSheets[sheets[entry.FirstSheetIndex].Name]), checked((ushort)nativeSheets[sheets[entry.LastSheetIndex].Name]));
                int existing = entries.FindIndex(item => selfBooks.Contains(item.SupBookIndex) && item.FirstSheetIndex == entry.FirstSheetIndex && item.LastSheetIndex == entry.LastSheetIndex);
                if (existing >= 0) indexMap.Add(checked((ushort)index), checked((ushort)existing)); else needsSelf = true;
            }
            nativeEntries.Add(entry);
        }
        var bookMap = new Dictionary<ushort, ushort>(); var links = new List<SupportingLinkRecord>(); int generatedBook = -1;
        foreach (var record in generated.SupportingLinkRecords) {
            if (record.RecordType == 0x01ae) {
                generatedBook++;
                if (generatedBook == 0 && !needsSelf) continue;
                bookMap.Add(checked((ushort)generatedBook), checked((ushort)(books.Length + bookMap.Count)));
            }
            byte[] payload = record.Payload;
            if (generatedBook == 0 && !needsSelf) continue;
            if (generatedBook == 0 && record.RecordType == 0x01ae) {
                payload = (byte[])payload.Clone(); payload[0] = (byte)nativeSheets.Count; payload[1] = checked((byte)(nativeSheets.Count >> 8));
            }
            links.Add(new SupportingLinkRecord(record.RecordType, payload));
        }
        for (int index = 0; index < nativeEntries.Count; index++) {
            ushort key = checked((ushort)index); if (indexMap.ContainsKey(key)) continue;
            var entry = nativeEntries[index];
            indexMap.Add(key, checked((ushort)entries.Count));
            entries.Add(new ExternSheetEntry(bookMap[entry.SupBookIndex], entry.FirstSheetIndex, entry.LastSheetIndex));
        }
        byte[] combined = BuildPayload(entries);
        if (combined.Length > 8224) throw new NotSupportedException("The combined native external-sheet table requires BIFF continuation encoding.");
        var result = new LegacyXlsExternSheetTable(generated.SheetNames,
            generated._rangeIndexes.ToDictionary(item => item.Key, item => indexMap[item.Value], StringComparer.Ordinal),
            generated._externalSheetIndexes.ToDictionary(item => item.Key, item => indexMap[item.Value], StringComparer.OrdinalIgnoreCase),
            generated._externalNameIndexes.ToDictionary(item => item.Key, item => new ExternalNameIndex(indexMap[item.Value.ExternSheetIndex], item.Value.OneBasedNameIndex), StringComparer.OrdinalIgnoreCase),
            links, combined);
        result._preservedSheetIndexes = Enumerable.Range(0, sheets.Count).Select(index => indexMap[checked((ushort)index)]).ToArray();
        return result;
    }

    private static ExternSheetEntry ReadEntry(byte[] bytes, int index) {
        int offset = checked(2 + index * 6);
        return new ExternSheetEntry(BiffRecordReader.ReadUInt16(bytes, offset), BiffRecordReader.ReadUInt16(bytes, offset + 2), BiffRecordReader.ReadUInt16(bytes, offset + 4));
    }
}
