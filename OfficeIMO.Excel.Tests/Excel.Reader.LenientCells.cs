using System.Data.Common;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

// Excel opens cells with an empty style attribute after a repair prompt and cells with an
// unparseable reference silently. These tests pin the readers to the same results.
public partial class Excel {
    private const string EmptyStyleRows =
        "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Amount</t></is></c></row>"
        + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\" s=\"\"><is><t>Alpha</t></is></c><c r=\"B2\" s=\"\"><v>12.5</v></c></row>";

    private const string ValidStyleRows =
        "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Amount</t></is></c></row>"
        + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\" s=\"0\"><is><t>Alpha</t></is></c><c r=\"B2\" s=\"0\"><v>12.5</v></c></row>";

    // Plain numeric cells keep the loaded worksheet on the grid loader, which records the repair.
    private const string EmptyStyleNumberRows =
        "<row r=\"1\"><c r=\"A1\" s=\"\"><v>1</v></c><c r=\"B1\" s=\"\"><v>2</v></c></row>";

    private const string UnparseableReferenceRows =
        "<row r=\"1\"><c r=\"A\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B1X\" t=\"inlineStr\"><is><t>Amount</t></is></c></row>"
        + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>Alpha</t></is></c><c r=\"D\"><v>12.5</v></c></row>";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DataReader_ReadsCellsWithEmptyStyleAttribute(bool utf16) {
        string path = CreateLenientCellWorkbook(EmptyStyleRows, utf16);
        try {
            AssertLenientDataReaderRow(path, new ExcelReaderOptions());
        } finally { File.Delete(path); }
    }

    [Fact]
    public void DataReader_ReportsEmptyStyleAttributeRepair() {
        string path = CreateLenientCellWorkbook(EmptyStyleRows);
        try {
            var messages = new List<string>();
            var options = new ExcelReaderOptions();
            options.Execution.OnInfo = messages.Add;

            AssertLenientDataReaderRow(path, options);

            string message = Assert.Single(messages, m => m.StartsWith("[ReadRepair EmptyCellStyle]", StringComparison.Ordinal));
            Assert.Contains("'Data' cell A2", message);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void DataReader_ReportsNoRepairForValidStyles() {
        string path = CreateLenientCellWorkbook(ValidStyleRows);
        try {
            var messages = new List<string>();
            var options = new ExcelReaderOptions();
            options.Execution.OnInfo = messages.Add;

            AssertLenientDataReaderRow(path, options);

            Assert.DoesNotContain(messages, m => m.StartsWith("[ReadRepair", StringComparison.Ordinal));
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Reader_ReadRange_ReadsCellsWithEmptyStyleAttribute(bool utf16) {
        string path = CreateLenientCellWorkbook(EmptyStyleRows, utf16);
        try {
            using var reader = ExcelDocumentReader.Open(path);
            object?[,] values = reader.GetSheet("Data").ReadRange("A2:B2", ExcelExecutionMode.Sequential);

            Assert.Equal("Alpha", values[0, 0]);
            Assert.Equal(12.5d, Convert.ToDouble(values[0, 1]));
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Reader_ReadRows_ReadsCellsWithEmptyStyleAttribute(bool utf16) {
        string path = CreateLenientCellWorkbook(EmptyStyleRows, utf16);
        try {
            using var reader = ExcelDocumentReader.Open(path);
            object?[] row = Assert.Single(reader.GetSheet("Data").ReadRows("A2:B2"))!;

            Assert.Equal("Alpha", row[0]);
            Assert.Equal(12.5d, Convert.ToDouble(row[1]));
        } finally { File.Delete(path); }
    }

    [Fact]
    public void Load_ReportsEmptyStyleAttributeRepair() {
        string path = CreateLenientCellWorkbook(EmptyStyleNumberRows);
        try {
            using var document = ExcelDocument.Load(path, null, out OfficeConversionReport report);
            Assert.True(document.GetSheet("Data").TryGetCellText(1, 1, out _));

            OfficeDiagnostic diagnostic = Assert.Single(report.Diagnostics);
            Assert.Equal(OfficeDiagnosticCodes.ExcelCellStyleRepaired, diagnostic.Code);
            Assert.Equal(OfficeDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Data!A1", diagnostic.Location);
            Assert.Contains("2 cells", diagnostic.Message);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void Load_ReportPicksUpRepairWhenTheWorksheetLoads() {
        string path = CreateLenientCellWorkbook(EmptyStyleNumberRows);
        try {
            using var document = ExcelDocument.Load(path, null, out OfficeConversionReport report);
            Assert.Empty(report.Diagnostics);

            Assert.True(document.GetSheet("Data").TryGetCellText(1, 1, out _));

            Assert.Equal(OfficeDiagnosticCodes.ExcelCellStyleRepaired, Assert.Single(report.Diagnostics).Code);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void Load_ReadsCellsWithEmptyStyleAttribute() {
        string path = CreateLenientCellWorkbook(EmptyStyleRows);
        try {
            using var document = ExcelDocument.Load(path);
            ExcelSheet sheet = document.GetSheet("Data");

            Assert.True(sheet.TryGetCellText(2, 1, out string name));
            Assert.Equal("Alpha", name);
            Assert.True(sheet.TryGetCellText(2, 2, out string amount));
            Assert.Equal("12.5", amount);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void Load_ReportsNoRepairForValidStyles() {
        string path = CreateLenientCellWorkbook(ValidStyleRows);
        try {
            using var document = ExcelDocument.Load(path, null, out OfficeConversionReport report);
            Assert.True(document.GetSheet("Data").TryGetCellText(2, 1, out _));

            Assert.Empty(report.Diagnostics);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void ValidateDocument_RejectsEmptyStyleAttribute() {
        string path = CreateLenientCellWorkbook(EmptyStyleRows);
        try {
            using var document = ExcelDocument.Load(path);

            Assert.Contains(document.ValidateDocument(), error => error.Description.Contains("'s'", StringComparison.Ordinal));
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DataReader_PlacesUnparseableCellReferencesInNextColumn(bool utf16) {
        string path = CreateLenientCellWorkbook(UnparseableReferenceRows, utf16);
        try {
            AssertLenientDataReaderRow(path, new ExcelReaderOptions());
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("A", "B1X")]
    [InlineData("A1", "D")]
    [InlineData("0", "1")]
    [InlineData("A1", "B")]
    public void Reader_ReadRange_PlacesUnparseableCellReferencesInNextColumn(string first, string second) {
        string path = CreateLenientCellWorkbook(
            $"<row r=\"1\"><c r=\"{first}\"><v>1</v></c><c r=\"{second}\"><v>2</v></c></row>");
        try {
            using var reader = ExcelDocumentReader.Open(path);
            object?[,] values = reader.GetSheet("Data").ReadRange("A1:B1", ExcelExecutionMode.Sequential);

            Assert.Equal(1d, Convert.ToDouble(values[0, 0]));
            Assert.Equal(2d, Convert.ToDouble(values[0, 1]));
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Reader_ReadRows_PlacesUnparseableCellReferencesInNextColumn(bool utf16) {
        string path = CreateLenientCellWorkbook(UnparseableReferenceRows, utf16);
        try {
            using var reader = ExcelDocumentReader.Open(path);
            var rows = reader.GetSheet("Data").ReadRows("A1:B2").ToList();

            Assert.Equal(new object?[] { "Name", "Amount" }, rows[0]);
            Assert.Equal("Alpha", rows[1]![0]);
            Assert.Equal(12.5d, Convert.ToDouble(rows[1]![1]));
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Load_PlacesUnparseableCellReferencesInNextColumn(bool utf16) {
        string path = CreateLenientCellWorkbook(UnparseableReferenceRows, utf16);
        try {
            using var document = ExcelDocument.Load(path);
            ExcelSheet sheet = document.GetSheet("Data");

            Assert.True(sheet.TryGetCellText(1, 1, out string name));
            Assert.Equal("Name", name);
            Assert.True(sheet.TryGetCellText(1, 2, out string amountHeader));
            Assert.Equal("Amount", amountHeader);
            Assert.True(sheet.TryGetCellText(2, 2, out string amount));
            Assert.Equal("12.5", amount);
        } finally { File.Delete(path); }
    }

    /// <summary>
    /// Single-pass reading indexes a late lenient cell instead of stopping after the rows it returned.
    /// </summary>
    [Theory]
    [InlineData("<c r=\"B23990\" s=\"1\">", "<c r=\"B23990\" s=\"\">")]
    [InlineData("<c r=\"C23990\">", "<c r=\"C\">")]
    [InlineData("<c r=\"C23990\">", "<c r=\" C&#50;3990 \">")]
    [InlineData("<c r=\"C23990\">", "<c r=\"é\">")]
    public void SinglePassReading_ReadsLateLenientCellsAsTheDefaultReaderDoes(string original, string replacement) {
        byte[] workbook = CreateParallelIndexWorkbook(invalidSharedStringRow: 0);
        using (var buffer = new MemoryStream()) {
            buffer.Write(workbook, 0, workbook.Length);
            using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true)) {
                string xml;
                using (var reader = new StreamReader(archive.GetEntry("xl/worksheets/sheet1.xml")!.Open())) xml = reader.ReadToEnd();
                Assert.Contains(original, xml);
                ReplaceAsyncCultureEntry(archive, "xl/worksheets/sheet1.xml", xml.Replace(original, replacement));
            }
            workbook = buffer.ToArray();
        }

        using DbDataReader expected = OpenParallelIndexReader(workbook, numericAsDecimal: true, sequential: true);
        using DbDataReader actual = ExcelDocument.OpenReader(new MemoryStream(workbook, writable: false),
            new ExcelReaderOptions { NumericAsDecimal = true, EnableSinglePassReading = true });
        while (expected.Read()) {
            Assert.True(actual.Read());
            for (int ordinal = 0; ordinal < expected.FieldCount; ordinal++) {
                Assert.Equal(expected.GetValue(ordinal), actual.GetValue(ordinal));
            }
        }
        Assert.False(actual.Read());
    }

    [Fact]
    public void Cell_PlacesUnparseableReferenceAfterStoredCell() {
        // The loader keeps B1 outside the DOM and the malformed cell in it; the malformed cell still belongs at C1.
        string path = CreateLenientCellWorkbook("<row r=\"1\"><c r=\"B1\"><v>1</v></c><c r=\"D\"><v>2</v></c></row>");
        try {
            using var document = ExcelDocument.Load(path);
            ExcelSheet sheet = document.GetSheet("Data");

            Assert.Equal(ExcelCellDataKind.Blank, sheet.Cell("A1").GetValue().Kind);
            Assert.Equal(1d, sheet.Cell("B1").GetValue().Value);
            Assert.Equal(2d, sheet.Cell("C1").GetValue().Value);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void Cell_WriteUpdatesTheCellReadAtAnUnparseableReference() {
        string path = CreateLenientCellWorkbook("<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"D\"><v>2</v></c></row>");
        try {
            using (var document = ExcelDocument.Load(path)) {
                ExcelSheet sheet = document.GetSheet("Data");
                sheet.Cell(1, 2).SetValue(5);

                Assert.Equal(5d, sheet.Cell("B1").GetValue().Value);
                Assert.Equal(ExcelCellDataKind.Blank, sheet.Cell("C1").GetValue().Kind);
                document.Save();
            }

            using var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, false);
            var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet!
                .Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Select(c => c.CellReference?.Value).ToArray();
            Assert.Equal(new[] { "A1", "B1" }, cells);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void SetValues_UpdatesTheCellReadAtAnUnparseableReference() {
        string path = CreateLenientCellWorkbook("<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"D\"><v>2</v></c></row>");
        try {
            using (var document = ExcelDocument.Load(path)) {
                ExcelSheet sheet = document.GetSheet("Data");
                // More cells than the per-cell write limit, on an existing row, take the prepared batch writer.
                var cells = new List<(int Row, int Column, object Value)> { (1, 2, 5) };
                for (int column = 1; column <= 16; column++) cells.Add((2, column, column));
                sheet.SetValues(cells, ExcelExecutionMode.Sequential);

                Assert.Equal(5d, sheet.Cell("B1").GetValue().Value);
                Assert.Equal(ExcelCellDataKind.Blank, sheet.Cell("C1").GetValue().Kind);
                document.Save();
            }

            using var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, false);
            var firstRow = package.WorkbookPart!.WorksheetParts.Single().Worksheet!
                .Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().First(row => row.RowIndex?.Value == 1);
            Assert.Equal(new[] { "A1", "B1" }, firstRow.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Select(c => c.CellReference?.Value).ToArray());
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Write_PlacesReferenceWithTrailingTextByTheReadRule(bool batch) {
        // Readers put B1X first in the row, at A1, so writing A1 updates it.
        string path = CreateLenientCellWorkbook("<row r=\"1\"><c r=\"B1X\"><v>1</v></c></row>");
        try {
            using (var document = ExcelDocument.Load(path)) {
                ExcelSheet sheet = document.GetSheet("Data");
                WriteLenientCell(sheet, 1, 1, 5, batch);

                Assert.Equal(5d, sheet.Cell("A1").GetValue().Value);
                Assert.Equal(ExcelCellDataKind.Blank, sheet.Cell("B1").GetValue().Kind);
                document.Save();
            }

            Assert.Equal(new[] { "A1" }, ReadSavedRowReferences(path, 1));
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SetValues_SeedsImplicitColumnFromThePrecedingCell(bool utf16) {
        // Readers place the cell without a reference after A1, at B1, not after the largest column.
        string path = CreateLenientCellWorkbook(
            "<row r=\"1\"><c r=\"C1\"><v>3</v></c><c r=\"A1\"><v>1</v></c><c><v>2</v></c></row>", utf16);
        try {
            using var document = ExcelDocument.Load(path);
            ExcelSheet sheet = document.GetSheet("Data");
            Assert.Equal(2d, sheet.Cell("B1").GetValue().Value);

            WriteLenientCell(sheet, 1, 4, 9, batch: true);

            Assert.Equal(2d, sheet.Cell("B1").GetValue().Value);
            Assert.Equal(3d, sheet.Cell("C1").GetValue().Value);
            Assert.Equal(9d, sheet.Cell("D1").GetValue().Value);
        } finally { File.Delete(path); }
    }

    private static void WriteLenientCell(ExcelSheet sheet, int row, int column, object value, bool batch) {
        if (!batch) {
            sheet.Cell(row, column).SetValue(value);
            return;
        }
        // More cells than the per-cell write limit, on an existing row, take the prepared batch writer.
        var cells = new List<(int Row, int Column, object Value)> { (row, column, value) };
        for (int other = 1; other <= 16; other++) cells.Add((row + 1, other, other));
        sheet.SetValues(cells, ExcelExecutionMode.Sequential);
    }

    private static string?[] ReadSavedRowReferences(string path, uint rowIndex) {
        using var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, false);
        return package.WorkbookPart!.WorksheetParts.Single().Worksheet!
            .Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().First(row => row.RowIndex?.Value == rowIndex)
            .Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Select(c => c.CellReference?.Value).ToArray();
    }

    [Theory]
    [InlineData("&#66;2")]
    [InlineData(" B2 ")]
    public void DataReader_ReadsReferencesTheByteParserDoesNotAccept(string reference) {
        string path = CreateLenientCellWorkbook(
            "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Amount</t></is></c></row>"
            + $"<row r=\"2\"><c r=\"{reference}\"><v>42</v></c></row>");
        try {
            using var reader = ExcelDocument.OpenReader(path);
            Assert.True(reader.Read());

            Assert.Equal(DBNull.Value, reader["Name"]);
            Assert.Equal(42d, Convert.ToDouble(reader["Amount"]));
        } finally { File.Delete(path); }
    }

    [Fact]
    public void Load_ReportDiagnosticsAreReadOnly() {
        string path = CreateLenientCellWorkbook(EmptyStyleNumberRows);
        try {
            using var document = ExcelDocument.Load(path, null, out OfficeConversionReport report);
            Assert.True(document.GetSheet("Data").TryGetCellText(1, 1, out _));

            var findings = Assert.IsAssignableFrom<IList<OfficeDiagnostic>>(report.Diagnostics);
            OfficeDiagnostic finding = Assert.Single(findings);
            Assert.Throws<NotSupportedException>(() => findings[0] = finding);
        } finally { File.Delete(path); }
    }

    private static void AssertLenientDataReaderRow(string path, ExcelReaderOptions options) {
        using var reader = ExcelDocument.OpenReader(path, options);
        Assert.True(reader.Read());
        Assert.Equal("Alpha", reader["Name"]);
        Assert.Equal(12.5d, Convert.ToDouble(reader["Amount"]));
        Assert.False(reader.Read());
    }

    private static string CreateLenientCellWorkbook(string rows, bool utf16 = false) {
        string path = CreateCompactFastPathWorkbook();
        string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
            + rows + "</sheetData></worksheet>";
        // UTF-16 parts skip the UTF-8 byte readers and take the XmlReader paths.
        ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", utf16
            ? Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray()
            : Encoding.UTF8.GetBytes(xml));
        return path;
    }
}
