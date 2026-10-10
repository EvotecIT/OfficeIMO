using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Write;
using System.Threading;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void Xlsb_NewUnformattedTextWorkbookDefinesDefaultStyle(bool multipleSheets, bool sharedStrings) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            if (multipleSheets) {
                sheet.CellValue(1, 1, "Text");
                sheet.CellValue(2, 1, "value");
                document.AddWorksheet("Other").CellValue(1, 1, "other");
            } else {
                sheet.InsertObjects(new[] { "value" }, includeHeaders: true, startRow: 1,
                    ("Text", static value => value));
            }

            byte[] workbook = document.ToBytes(ExcelFileFormat.Xlsb,
                new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings });

            if (multipleSheets) Assert.NotEqual(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            else Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            AssertNativeDefaultStyle(workbook);

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook);
            Assert.Equal("Text", reader.GetName(0));
            Assert.True(reader.Read());
            Assert.Equal("value", reader.GetString(0));
            Assert.False(reader.Read());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Xlsb_StagedUnformattedWorkbookDefinesDefaultStyle(bool sharedStrings) {
            var rows = new SnapshotXlsbRows(aboveCaptureLimit: true);
            rows.Values[0][2] = null;
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data");
            var source = new ExcelDirectTabularSource("Data", rows, includeHeaders: true, preserveMissingValues: true);
            using var destination = new MemoryStream();

            Assert.True(XlsbNewPackageWriter.TryWriteDirectTabular(
                document, source, destination, CancellationToken.None, sharedStrings));

            byte[] workbook = destination.ToArray();
            AssertNativeDefaultStyle(workbook);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook);
            Assert.Equal("Text", reader.GetName(0));
            Assert.True(reader.Read());
            Assert.Equal("Zażółć 🚀", reader.GetString(0));
            Assert.Equal(42D, reader.GetDouble(1));
            Assert.True(reader.IsDBNull(2));
            Assert.True(reader.GetBoolean(3));
        }

        private static void AssertNativeDefaultStyle(byte[] workbook) {
            using var styles = new MemoryStream(ReadPackageEntry(workbook, "xl/styles.bin"), writable: false);
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(styles);
            XlsbRecord cellFormats = Assert.Single(records, record => record.Type == 617);
            Assert.Equal(new byte[] { 1, 0, 0, 0 }, cellFormats.Data);
            Assert.Contains(records, record => record.Type == 47);

            XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            using var relationshipsStream = new MemoryStream(ReadPackageEntry(workbook, "xl/_rels/workbook.bin.rels"), writable: false);
            XElement link = Assert.Single(XDocument.Load(relationshipsStream).Root!.Elements(relationships + "Relationship"),
                element => (string?)element.Attribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            Assert.Equal("styles.bin", (string?)link.Attribute("Target"));
        }
    }
}
