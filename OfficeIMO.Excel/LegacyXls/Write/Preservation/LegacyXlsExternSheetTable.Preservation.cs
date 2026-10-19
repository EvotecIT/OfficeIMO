using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal sealed partial class LegacyXlsExternSheetTable {
        internal static LegacyXlsExternSheetTable CreatePreserving(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets,
            IReadOnlyList<OfficeIMO.Excel.LegacyXls.Biff.BiffRecord> source, IReadOnlyDictionary<string, int> nativeSheets) {
            var generated = Create(document, sheets);
            var old = source.Where(record => record.Type == 0x0017).ToArray();
            if (old.Length > 1) throw new NotSupportedException("Multiple native ExternSheet records require consolidation.");
            byte[] original = old.Length == 0 ? new byte[2] : old[0].Payload;
            int count = OfficeIMO.Excel.LegacyXls.Biff.BiffRecordReader.ReadUInt16(original, 0);
            if (original.Length != 2 + count * 6) throw new NotSupportedException("Invalid or continued native ExternSheet table.");
            int generatedCount = OfficeIMO.Excel.LegacyXls.Biff.BiffRecordReader.ReadUInt16(generated.Payload, 0);
            int supBooks = source.Count(record => record.Type == 0x01ae);
            byte[] combined = new byte[checked(original.Length + generatedCount * 6)];
            Buffer.BlockCopy(original, 0, combined, 0, original.Length);
            void Put(int offset, int value) { combined[offset] = checked((byte)(value & 255)); combined[offset + 1] = checked((byte)(value >> 8)); }
            Put(0, checked(count + generatedCount));
            for (int index = 0; index < generatedCount; index++) {
                int from = 2 + index * 6, to = original.Length + index * 6;
                int book = OfficeIMO.Excel.LegacyXls.Biff.BiffRecordReader.ReadUInt16(generated.Payload, from);
                int first = OfficeIMO.Excel.LegacyXls.Biff.BiffRecordReader.ReadUInt16(generated.Payload, from + 2);
                int last = OfficeIMO.Excel.LegacyXls.Biff.BiffRecordReader.ReadUInt16(generated.Payload, from + 4);
                Put(to, checked(book + supBooks));
                Put(to + 2, book == 0 && first < sheets.Count ? nativeSheets[sheets[first].Name] : first);
                Put(to + 4, book == 0 && last < sheets.Count ? nativeSheets[sheets[last].Name] : last);
            }
            var links = generated.SupportingLinkRecords.Select(record => {
                byte[] payload = (byte[])record.Payload.Clone();
                if (record.RecordType == 0x01ae && payload.Length == 4 && payload[2] == 1 && payload[3] == 4) {
                    payload[0] = checked((byte)(nativeSheets.Count & 255)); payload[1] = checked((byte)(nativeSheets.Count >> 8));
                }
                return new SupportingLinkRecord(record.RecordType, payload);
            }).ToArray();
            var result = new LegacyXlsExternSheetTable(generated.SheetNames,
                generated._rangeIndexes.ToDictionary(item => item.Key, item => checked((ushort)(item.Value + count)), StringComparer.Ordinal),
                generated._externalSheetIndexes.ToDictionary(item => item.Key, item => checked((ushort)(item.Value + count)), StringComparer.OrdinalIgnoreCase),
                generated._externalNameIndexes.ToDictionary(item => item.Key, item => new ExternalNameIndex(checked((ushort)(item.Value.ExternSheetIndex + count)), item.Value.OneBasedNameIndex), StringComparer.OrdinalIgnoreCase),
                links, combined);
            result._nativeIndexOffset = checked((ushort)count);
            return result;
        }

    }
}
