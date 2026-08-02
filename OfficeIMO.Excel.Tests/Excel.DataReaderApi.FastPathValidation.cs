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
                "r=\"A2\" r=\"A2\"");
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
                "<!--invalid--comment--></sheetData>");
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
                "<sheetData bad:attribute=\"value\">");
            Assert.NotEqual(worksheetXml, malformedXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RejectsMalformedXmlDeclarationOnCompactFastPath() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            int declarationEnd = worksheetXml.IndexOf("?>", StringComparison.Ordinal);
            Assert.StartsWith("<?xml", worksheetXml, StringComparison.Ordinal);
            Assert.True(declarationEnd > 0);
            string malformedXml = "<?xml?>" + worksheetXml.Substring(declarationEnd + 2);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RejectsAdjacentCellAttributesOnCompactFastPath() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformedXml = worksheetXml.Replace(
                "r=\"A2\" t=\"n\"",
                "r=\"A2\"t=\"n\"");
            Assert.NotEqual(worksheetXml, malformedXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("<row", "< row")]
    [InlineData("</row>", "</ row>")]
    public void OpenDataReader_RejectsWhitespaceInsideRowTagOpeners(
        string validToken,
        string malformedToken) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformedXml = worksheetXml.Replace(
                validToken,
                malformedToken);
            Assert.NotEqual(worksheetXml, malformedXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RejectsUnboundNamespacePrefixOnIndexedRow() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformedXml = worksheetXml.Replace(
                "<row r=\"2\"",
                "<row bad:x=\"1\" r=\"2\"");
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
