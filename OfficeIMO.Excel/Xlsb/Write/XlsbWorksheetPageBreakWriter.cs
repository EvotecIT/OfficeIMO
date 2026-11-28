using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Encodes worksheet row and column page breaks as RWBRK and COLBRK collections (MS-XLSB 2.4.197, 2.4.25, 2.4.315).</summary>
    internal static class XlsbWorksheetPageBreakWriter {
        internal const int BrtBeginRwBrk = 392;
        internal const int BrtEndRwBrk = 393;
        internal const int BrtBeginColBrk = 394;
        internal const int BrtEndColBrk = 395;
        internal const int BrtBrk = 396;
        private const int MaximumBreaks = 1023;
        private static readonly string[] CollectionAttributes = { "count", "manualBreakCount" };
        private static readonly string[] BreakAttributes = { "id", "min", "max", "man", "pt" };

        internal static void Append(XlsbWorksheetFeatureRecords records, XlsbWorksheetFeatureContext context) {
            records.PageBreaks.AddRange(CreateRowBreaks(context));
            records.PageBreaks.AddRange(CreateColumnBreaks(context));
        }

        internal static IReadOnlyList<XlsbGeneratedRecord> CreateRowBreaks(XlsbWorksheetFeatureContext context) =>
            Create(XlsbWorksheetMarkup.Child(context.WorksheetMarkup, "rowBreaks"), rows: true, context.Sheet.Name);

        internal static IReadOnlyList<XlsbGeneratedRecord> CreateColumnBreaks(XlsbWorksheetFeatureContext context) =>
            Create(XlsbWorksheetMarkup.Child(context.WorksheetMarkup, "colBreaks"), rows: false, context.Sheet.Name);

        private static IReadOnlyList<XlsbGeneratedRecord> Create(OfficeXmlNode collection, bool rows, string sheetName) {
            if (collection.IsNone) return Array.Empty<XlsbGeneratedRecord>();
            string kind = rows ? "row" : "column";
            List<OfficeXmlNode> breaks = collection.Children(OfficeXmlNamespaces.SpreadsheetML, "brk").ToList();
            if (collection.Children().Count != breaks.Count
                || XlsbWorkbookMarkup.UnsupportedAttribute(collection, CollectionAttributes) != null) {
                throw new NotSupportedException($"Native XLSB generation found unsupported {kind} page-break content on worksheet '{sheetName}'.");
            }
            // RWBRK and COLBRK require at least one BrtBrk; an empty collection has nothing to keep.
            if (breaks.Count == 0) return Array.Empty<XlsbGeneratedRecord>();
            if (breaks.Count > MaximumBreaks) {
                throw new NotSupportedException($"Native XLSB generation supports at most {MaximumBreaks} {kind} page breaks on worksheet '{sheetName}'.");
            }

            uint indexLimit = rows ? 1_048_576U : 16_384U;
            uint spanLimit = rows ? 16_384U : 1_048_576U;
            var result = new List<XlsbGeneratedRecord>(breaks.Count + 2);
            var begin = new byte[8];
            WriteUInt32(begin, 0, (uint)breaks.Count);
            // ibrkManMac must equal ibrkMac (MS-XLSB 2.4.197); fMan on each BrtBrk still says which were set by hand.
            WriteUInt32(begin, 4, (uint)breaks.Count);
            result.Add(new XlsbGeneratedRecord(rows ? BrtBeginRwBrk : BrtBeginColBrk, begin));
            foreach (OfficeXmlNode item in breaks) {
                if (item.FirstChild.Exists || XlsbWorkbookMarkup.UnsupportedAttribute(item, BreakAttributes) != null) {
                    throw new NotSupportedException($"Native XLSB generation found an unsupported {kind} page break on worksheet '{sheetName}'.");
                }
                uint id = XlsbWorkbookMarkup.ReadUInt32(item, "id") ?? 0U;
                uint min = XlsbWorkbookMarkup.ReadUInt32(item, "min") ?? 0U;
                uint max = XlsbWorkbookMarkup.ReadUInt32(item, "max") ?? 0U;
                if (id >= indexLimit || min >= spanLimit || max >= spanLimit || max < min) {
                    throw new NotSupportedException($"Native XLSB generation found an out-of-range {kind} page break on worksheet '{sheetName}'.");
                }
                var payload = new byte[20];
                WriteUInt32(payload, 0, id);
                WriteUInt32(payload, 4, min);
                WriteUInt32(payload, 8, max);
                WriteUInt32(payload, 12, XlsbWorkbookMarkup.ReadBoolean(item, "man") == true ? 1U : 0U);
                WriteUInt32(payload, 16, XlsbWorkbookMarkup.ReadBoolean(item, "pt") == true ? 1U : 0U);
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
