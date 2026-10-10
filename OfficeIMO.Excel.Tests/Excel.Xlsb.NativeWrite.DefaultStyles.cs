using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
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
            sheet.CellValue(1, 1, "Text");
            sheet.CellValue(2, 1, "value");
            if (multipleSheets) document.AddWorksheet("Other").CellValue(1, 1, "other");

            byte[] workbook = document.ToBytes(ExcelFileFormat.Xlsb,
                new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings });

            if (multipleSheets) Assert.NotEqual(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            else Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            using var styles = new MemoryStream(ReadPackageEntry(workbook, "xl/styles.bin"), writable: false);
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(styles);
            XlsbRecord cellFormats = Assert.Single(records.Where(record => record.Type == 617));
            Assert.Equal(new byte[] { 1, 0, 0, 0 }, cellFormats.Data);
            Assert.Contains(records, record => record.Type == 47);

            XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            using var relationshipsStream = new MemoryStream(ReadPackageEntry(workbook, "xl/_rels/workbook.bin.rels"), writable: false);
            XElement link = Assert.Single(XDocument.Load(relationshipsStream).Root!.Elements(relationships + "Relationship")
                .Where(element => (string?)element.Attribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"));
            Assert.Equal("styles.bin", (string?)link.Attribute("Target"));

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook);
            Assert.Equal("Text", reader.GetName(0));
            Assert.True(reader.Read());
            Assert.Equal("value", reader.GetString(0));
            Assert.False(reader.Read());
        }
    }
}
