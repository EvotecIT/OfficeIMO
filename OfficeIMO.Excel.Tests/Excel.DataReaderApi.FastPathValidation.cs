using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void OpenDataReader_ReadsCanonicalDenseRowsWhenValidatedRowShapeChanges() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string canonicalXml = worksheetXml.Replace(" t=\"n\"", string.Empty);
            string changedRowXml = canonicalXml.Replace(
                "<row r=\"3\"",
                "<row customFormat=\"1\" r=\"3\"");
            Assert.NotEqual(canonicalXml, changedRowXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(changedRowXml));

            AssertCompactNumericRows(path);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_ValidatesMarkupBetweenCanonicalDenseRows() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string canonicalXml = worksheetXml.Replace(" t=\"n\"", string.Empty);
            string validXml = canonicalXml.Replace(
                "<row r=\"3\"",
                "<!--validated boundary--><row r=\"3\"");
            Assert.NotEqual(canonicalXml, validXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(validXml));
            AssertCompactNumericRows(path);

            string malformedXml = canonicalXml.Replace(
                "<row r=\"3\"",
                "<!--invalid--comment--><row r=\"3\"");
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));
            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

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

    [Theory]
    [InlineData("UTF-16")]
    [InlineData("bogus")]
    public void OpenDataReader_RejectsIncompatibleXmlDeclarationEncodingOnCompactFastPath(
        string encodingName) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            const string encodingMarker = "encoding=\"";
            int valueStart = worksheetXml.IndexOf(encodingMarker, StringComparison.Ordinal)
                + encodingMarker.Length;
            Assert.True(valueStart >= encodingMarker.Length);
            int valueEnd = worksheetXml.IndexOf('"', valueStart);
            Assert.True(valueEnd > valueStart);
            string malformedXml = worksheetXml.Substring(0, valueStart)
                + encodingName
                + worksheetXml.Substring(valueEnd);
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

    [Theory]
    [InlineData("\ufffe")]
    [InlineData("\uffff")]
    public void OpenDataReader_RejectsXmlForbiddenScalarsOnIndexedRows(string forbiddenScalar) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformedXml = worksheetXml.Replace(
                "<row r=\"2\"",
                $"<row note=\"{forbiddenScalar}\" r=\"2\"");
            Assert.NotEqual(worksheetXml, malformedXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("http://www.w3.org/XML/1998/namespace")]
    [InlineData("http://www.w3.org/2000/xmlns/")]
    public void OpenDataReader_RejectsReservedNamespaceBindings(string namespaceUri) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformedXml = worksheetXml.Replace(
                "<worksheet ",
                $"<worksheet xmlns:p=\"{namespaceUri}\" ");
            Assert.NotEqual(worksheetXml, malformedXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RejectsRawCDataTerminatorInsideFastValidatedSheetData() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformedXml = worksheetXml.Replace(
                "<row r=\"2\"",
                "]]><row r=\"2\"");
            Assert.NotEqual(worksheetXml, malformedXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("p:")]
    [InlineData("p:1value")]
    [InlineData("p:value:extra")]
    public void OpenDataReader_RejectsMalformedPrefixedAttributeNames(string attributeName) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformedXml = worksheetXml
                .Replace(
                    "<worksheet ",
                    "<worksheet xmlns:p=\"urn:officeimo:test\" ")
                .Replace(
                    "<row r=\"2\"",
                    $"<row {attributeName}=\"value\" r=\"2\"");
            Assert.NotEqual(worksheetXml, malformedXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RejectsDuplicateExpandedNamespaceAttributes() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformedXml = worksheetXml
                .Replace(
                    "<worksheet ",
                    "<worksheet xmlns:p=\"urn:officeimo:test\" xmlns:q=\"urn:officeimo:test\" ")
                .Replace(
                    "<row r=\"2\"",
                    "<row p:value=\"one\" q:value=\"two\" r=\"2\"");
            Assert.NotEqual(worksheetXml, malformedXml);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformedXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RejectsCharacterDataBeforeWorksheetRoot() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            int rootStart = worksheetXml.IndexOf("<worksheet", StringComparison.Ordinal);
            Assert.True(rootStart > 0);
            string malformedXml = worksheetXml.Insert(rootStart, "junk");
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
            sheet.CellValue(3, 1, 43);
            document.Save();
        }
        return path;
    }

    private static void AssertCompactNumericRows(string path) {
        using var reader = ExcelDocument.OpenDataReader(path);
        Assert.True(reader.Read());
        Assert.Equal(42, reader.GetInt32(0));
        Assert.True(reader.Read());
        Assert.Equal(43, reader.GetInt32(0));
        Assert.False(reader.Read());
    }
}
