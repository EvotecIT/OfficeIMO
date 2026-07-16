using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Writes the workbook-level BIFF12 records for a newly generated XLSB package.</summary>
    internal static class XlsbWorkbookPartWriter {
        private const int BrtBeginBook = 131;
        private const int BrtEndBook = 132;
        private const int BrtBeginBundleShs = 143;
        private const int BrtEndBundleShs = 144;
        private const int BrtWbProp = 153;
        private const int BrtBundleSh = 156;

        internal static byte[] Create(
            IReadOnlyList<ExcelSheet> sheets,
            bool uses1904DateSystem,
            DocumentFormat.OpenXml.Spreadsheet.BookViews? workbookViews,
            DocumentFormat.OpenXml.Spreadsheet.WorkbookProtection? workbookProtection,
            DocumentFormat.OpenXml.Spreadsheet.DefinedNames? definedNames,
            DocumentFormat.OpenXml.Spreadsheet.CalculationProperties? calculationProperties) {
            if (sheets == null) throw new ArgumentNullException(nameof(sheets));
            if (sheets.Count == 0) throw new NotSupportedException("Native XLSB generation requires at least one worksheet.");

            using var output = new MemoryStream(Math.Max(256, sheets.Count * 64));
            XlsbRecordWriter.Write(output, BrtBeginBook);
            XlsbRecordWriter.Write(output, BrtWbProp, CreateWorkbookPropertiesPayload(uses1904DateSystem));
            XlsbWorkbookProtectionWriter.Write(output, workbookProtection);
            XlsbWorkbookViewWriter.Write(output, workbookViews, sheets.Count);
            XlsbRecordWriter.Write(output, BrtBeginBundleShs);
            for (int index = 0; index < sheets.Count; index++) {
                XlsbRecordWriter.Write(output, BrtBundleSh, CreateBundleSheetPayload(sheets[index], index));
            }
            XlsbRecordWriter.Write(output, BrtEndBundleShs);
            XlsbDefinedNameWriter.Write(output, definedNames, sheets);
            XlsbCalculationPropertiesWriter.Write(output, calculationProperties);
            XlsbRecordWriter.Write(output, BrtEndBook);
            return output.ToArray();
        }

        private static byte[] CreateWorkbookPropertiesPayload(bool uses1904DateSystem) {
            var payload = new byte[12];
            if (uses1904DateSystem) payload[0] = 0x01;
            return payload;
        }

        private static byte[] CreateBundleSheetPayload(ExcelSheet sheet, int index) {
            string relationshipId = "rId" + (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            using var payload = new MemoryStream(32 + (relationshipId.Length + sheet.Name.Length) * 2);
            WriteUInt32(payload, sheet.VeryHidden ? 2U : sheet.Hidden ? 1U : 0U);
            WriteUInt32(payload, checked((uint)(index + 1)));
            WriteWideString(payload, relationshipId);
            WriteWideString(payload, sheet.Name);
            return payload.ToArray();
        }

        private static void WriteWideString(Stream stream, string value) {
            WriteUInt32(stream, checked((uint)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 24));
        }
    }
}
