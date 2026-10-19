using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Biff;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal sealed partial class LegacyXlsStyleTable {
        internal static LegacyXlsStyleTable CreatePreserving(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets,
            LegacyXlsFontTable fonts, IReadOnlyList<BiffRecord> source) {
            var generated = Create(document, sheets, fonts);
            int nativeCount = source.Count(record => record.Type == 0x00e0);
            var nativeFormats = source.Where(record => record.Type == 0x041e).ToDictionary(
                record => BiffRecordReader.ReadUInt16(record.Payload, 0), record => record.Payload);
            var remap = new Dictionary<ushort, ushort>();
            var formats = new List<byte[]>();
            foreach (byte[] format in generated.FormatRecords) {
                ushort index = BiffRecordReader.ReadUInt16(format, 0);
                if (nativeFormats.TryGetValue(index, out byte[]? native)) {
                    if (native.SequenceEqual(format)) continue;
                    ushort replacement = 164;
                    while (nativeFormats.ContainsKey(replacement)) replacement = checked((ushort)(replacement + 1));
                    var renamed = (byte[])format.Clone(); renamed[0] = (byte)replacement; renamed[1] = (byte)(replacement >> 8);
                    remap.Add(index, replacement); nativeFormats.Add(replacement, renamed); formats.Add(renamed);
                } else { nativeFormats.Add(index, format); formats.Add(format); }
            }
            int required = BuildRequiredStyleXfPayloads().Count;
            var appended = generated.CellFormatRecords.Skip(required).Select(bytes => (byte[])bytes.Clone()).ToArray();
            foreach (byte[] xf in appended) {
                if (remap.TryGetValue(BiffRecordReader.ReadUInt16(xf, 2), out ushort replacement)) {
                    xf[2] = (byte)replacement; xf[3] = (byte)(replacement >> 8);
                }
            }
            var indexes = generated._styleIndexMap.ToDictionary(item => item.Key, item => checked((ushort)(nativeCount + item.Value - required)));
            return new LegacyXlsStyleTable(indexes, formats, appended, Array.Empty<byte[]>());
        }

    }
}
