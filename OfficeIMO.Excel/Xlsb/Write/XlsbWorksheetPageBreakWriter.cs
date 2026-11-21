using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Encodes worksheet row and column page breaks as RWBRK and COLBRK collections (MS-XLSB 2.4.197, 2.4.25, 2.4.315).</summary>
    internal static class XlsbWorksheetPageBreakWriter {
        internal const int BrtBeginRwBrk = 392;
        internal const int BrtEndRwBrk = 393;
        internal const int BrtBeginColBrk = 394;
        internal const int BrtEndColBrk = 395;
        internal const int BrtBrk = 396;
        private const int MaximumBreaks = 1023;

        internal static void Append(XlsbWorksheetFeatureRecords records, XlsbWorksheetFeatureContext context) {
            records.PageBreaks.AddRange(CreateRowBreaks(context));
            records.PageBreaks.AddRange(CreateColumnBreaks(context));
        }

        internal static IReadOnlyList<XlsbGeneratedRecord> CreateRowBreaks(XlsbWorksheetFeatureContext context) =>
            Create(context.Worksheet.GetFirstChild<RowBreaks>(), rows: true, context.Sheet.Name);

        internal static IReadOnlyList<XlsbGeneratedRecord> CreateColumnBreaks(XlsbWorksheetFeatureContext context) =>
            Create(context.Worksheet.GetFirstChild<ColumnBreaks>(), rows: false, context.Sheet.Name);

        private static IReadOnlyList<XlsbGeneratedRecord> Create(OpenXmlCompositeElement? collection, bool rows, string sheetName) {
            if (collection == null) return Array.Empty<XlsbGeneratedRecord>();
            string kind = rows ? "row" : "column";
            if (collection.ChildElements.Any(child => child is not Break)
                || XlsbOpenXmlAttributeValidator.FindUnsupported(collection, new[] { "count", "manualBreakCount" }) is OpenXmlAttribute) {
                throw new NotSupportedException($"Native XLSB generation found unsupported {kind} page-break content on worksheet '{sheetName}'.");
            }
            Break[] breaks = collection.Elements<Break>().ToArray();
            // RWBRK and COLBRK require at least one BrtBrk; an empty collection has nothing to keep.
            if (breaks.Length == 0) return Array.Empty<XlsbGeneratedRecord>();
            if (breaks.Length > MaximumBreaks) {
                throw new NotSupportedException($"Native XLSB generation supports at most {MaximumBreaks} {kind} page breaks on worksheet '{sheetName}'.");
            }

            uint indexLimit = rows ? 1_048_576U : 16_384U;
            uint spanLimit = rows ? 16_384U : 1_048_576U;
            var result = new List<XlsbGeneratedRecord>(breaks.Length + 2);
            var begin = new byte[8];
            WriteUInt32(begin, 0, (uint)breaks.Length);
            // ibrkManMac must equal ibrkMac (MS-XLSB 2.4.197); fMan on each BrtBrk still says which were set by hand.
            WriteUInt32(begin, 4, (uint)breaks.Length);
            result.Add(new XlsbGeneratedRecord(rows ? BrtBeginRwBrk : BrtBeginColBrk, begin));
            foreach (Break item in breaks) {
                if (item.HasChildren
                    || XlsbOpenXmlAttributeValidator.FindUnsupported(item, new[] { "id", "min", "max", "man", "pt" }) is OpenXmlAttribute) {
                    throw new NotSupportedException($"Native XLSB generation found an unsupported {kind} page break on worksheet '{sheetName}'.");
                }
                uint id = item.Id?.Value ?? 0U;
                uint min = item.Min?.Value ?? 0U;
                uint max = item.Max?.Value ?? 0U;
                if (id >= indexLimit || min >= spanLimit || max >= spanLimit || max < min) {
                    throw new NotSupportedException($"Native XLSB generation found an out-of-range {kind} page break on worksheet '{sheetName}'.");
                }
                var payload = new byte[20];
                WriteUInt32(payload, 0, id);
                WriteUInt32(payload, 4, min);
                WriteUInt32(payload, 8, max);
                WriteUInt32(payload, 12, item.ManualPageBreak?.Value == true ? 1U : 0U);
                WriteUInt32(payload, 16, item.PivotTablePageBreak?.Value == true ? 1U : 0U);
                result.Add(new XlsbGeneratedRecord(BrtBrk, payload));
            }
            result.Add(new XlsbGeneratedRecord(rows ? BrtEndRwBrk : BrtEndColBrk, Array.Empty<byte>()));
            return result;
        }

        private static void WriteUInt32(byte[] buffer, int offset, uint value) {
            buffer[offset] = (byte)value;
            buffer[offset + 1] = (byte)(value >> 8);
            buffer[offset + 2] = (byte)(value >> 16);
            buffer[offset + 3] = (byte)(value >> 24);
        }
    }
}
