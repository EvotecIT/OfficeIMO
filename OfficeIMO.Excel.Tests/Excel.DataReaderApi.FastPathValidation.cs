using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void OpenDataReader_RejectsDuplicateCellAttributesOnCompactFastPath() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformedXml = worksheetXml.Replace(
                "r=\"A2\"",
                "r=\"A2\" r=\"A2\"",
                StringComparison.Ordinal);
            Assert.NotEqual(worksheetXml, malformedXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RejectsMalformedCommentInsideFastValidatedSheetData() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformedXml = worksheetXml.Replace(
                "</sheetData>",
                "<!--invalid--comment--></sheetData>",
                StringComparison.Ordinal);
            Assert.NotEqual(worksheetXml, malformedXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RejectsUnboundNamespacePrefixOutsideSheetData() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformedXml = worksheetXml.Replace(
                "<sheetData>",
                "<sheetData bad:attribute=\"value\">",
                StringComparison.Ordinal);
            Assert.NotEqual(worksheetXml, malformedXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    private static string CreateCompactFastPathWorkbook() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.FastPathValidation.{Guid.NewGuid():N}.xlsx");
        using (var document = ExcelDocument.Create(path)) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Value");
            sheet.CellValue(2, 1, 42);
            document.Save();
        }
        return path;
    }
}
