using System.Data.Common;
using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    private const int ParallelIndexRows = 24_000;

    /// <summary>
    /// Large uniform worksheets are indexed in background chunks, including attribute-free rows
    /// without a declared dimension. Irregular rows, a late formula and whitespace between rows
    /// must still produce exactly the sequential reader's values.
    /// </summary>
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void LargeUniformWorksheetReadsMatchSequentialIndexing(bool numericAsDecimal, bool implicitLayout) {
        byte[] workbook = CreateParallelIndexWorkbook(invalidSharedStringRow: 0, implicitLayout);
        using DbDataReader expected = OpenParallelIndexReader(workbook, numericAsDecimal, sequential: true);
        using DbDataReader actual = OpenParallelIndexReader(workbook, numericAsDecimal, sequential: false);
        int rows = 0;
        while (expected.Read()) {
            Assert.True(actual.Read());
            rows++;
            Assert.Equal(expected.GetString(0), actual.GetString(0));
            Assert.Equal(expected.GetDateTime(1), actual.GetDateTime(1));
            Assert.Equal(expected.IsDBNull(2), actual.IsDBNull(2));
            if (!expected.IsDBNull(2)) Assert.Equal(expected.GetInt32(2), actual.GetInt32(2));
            for (int ordinal = 3; ordinal <= 5; ordinal++) {
                Assert.Equal(expected.GetDecimal(ordinal), actual.GetDecimal(ordinal));
                Assert.Equal(expected.GetDouble(ordinal), actual.GetDouble(ordinal));
                Assert.Equal(expected.GetValue(ordinal), actual.GetValue(ordinal));
            }
            Assert.Equal(expected.GetValue(6), actual.GetValue(6));
        }
        Assert.False(actual.Read());
        Assert.Equal(ParallelIndexRows - 1, rows);
    }

    /// <summary>
    /// Shapes written by common producers: a namespace prefix, cells with both style and type,
    /// xml:space on text and on cells, and indentation between rows and cells.
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ProducerShapedWorksheetsReadMatchSequentialIndexing(bool prefixed) {
        byte[] workbook = CreateProducerShapedWorkbook(prefixed);
        using DbDataReader expected = OpenParallelIndexReader(workbook, numericAsDecimal: false, sequential: true);
        using DbDataReader actual = OpenParallelIndexReader(workbook, numericAsDecimal: false, sequential: false);
        int rows = 0;
        while (expected.Read()) {
            Assert.True(actual.Read());
            rows++;
            for (int ordinal = 0; ordinal < expected.FieldCount; ordinal++) {
                Assert.Equal(expected.GetValue(ordinal), actual.GetValue(ordinal));
            }
        }
        Assert.False(actual.Read());
        Assert.Equal(ParallelIndexRows - 1, rows);
    }

    [Fact]
    public void LargeUniformWorksheetReportsLateInvalidCellBeforeTheFirstRow() {
        byte[] workbook = CreateParallelIndexWorkbook(invalidSharedStringRow: ParallelIndexRows - 3);
        InvalidDataException sequential = Assert.Throws<InvalidDataException>(() => ReadFirstParallelIndexRow(workbook, sequential: true));
        InvalidDataException parallel = Assert.Throws<InvalidDataException>(() => ReadFirstParallelIndexRow(workbook, sequential: false));
        Assert.Contains($"cell A{ParallelIndexRows - 3} references a missing shared string", parallel.Message);
        Assert.Equal(sequential.Message, parallel.Message);
    }

    private static void ReadFirstParallelIndexRow(byte[] workbook, bool sequential) {
        using DbDataReader reader = OpenParallelIndexReader(workbook, numericAsDecimal: true, sequential);
        reader.Read();
    }

    private static DbDataReader OpenParallelIndexReader(byte[] workbook, bool numericAsDecimal, bool sequential) {
        var options = new ExcelReaderOptions { NumericAsDecimal = numericAsDecimal };
        if (sequential) options.Execution.Mode = ExcelExecutionMode.Sequential;
        return ExcelDocument.OpenReader(new MemoryStream(workbook, writable: false), options);
    }

    private static byte[] CreateProducerShapedWorkbook(bool prefixed) {
        string p = prefixed ? "x:" : string.Empty;
        string ns = prefixed ? " xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"" : " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";
        var sheet = new StringBuilder(ParallelIndexRows * 400);
        sheet.Append('<').Append(p).Append("worksheet").Append(ns).Append("><").Append(p).Append("dimension ref=\"A1:D")
            .Append(ParallelIndexRows).Append("\"/><").Append(p).Append("sheetData>\n");
        sheet.Append('<').Append(p).Append("row r=\"1\">");
        foreach (string header in new[] { "Region", "Note", "Date", "Amount" }) {
            sheet.Append('<').Append(p).Append("c t=\"inlineStr\"><").Append(p).Append("is><").Append(p).Append("t>")
                .Append(header).Append("</").Append(p).Append("t></").Append(p).Append("is></").Append(p).Append("c>");
        }
        sheet.Append("</").Append(p).Append("row>\n");
        for (int row = 2; row <= ParallelIndexRows; row++) {
            sheet.Append("  <").Append(p).Append("row r=\"").Append(row).Append("\">\n    ");
            sheet.Append('<').Append(p).Append("c r=\"A").Append(row).Append("\" s=\"0\" t=\"s\"><").Append(p).Append("v>")
                .Append(1 + row % 7).Append("</").Append(p).Append("v></").Append(p).Append("c>\n    ");
            string text = row % 3 == 0 ? "note " + row % 11 + " " : "note" + row % 11;
            string space = prefixed || row % 3 != 0 ? string.Empty : " xml:space=\"preserve\"";
            sheet.Append('<').Append(p).Append("c r=\"B").Append(row).Append("\" t=\"inlineStr\"><").Append(p).Append("is><").Append(p)
                .Append('t').Append(space).Append('>').Append(text).Append("</").Append(p).Append("t></").Append(p).Append("is></").Append(p).Append("c>\n    ");
            sheet.Append('<').Append(p).Append("c r=\"C").Append(row).Append("\" s=\"1\"><").Append(p).Append("v>")
                .Append(40000 + row % 900).Append("</").Append(p).Append("v></").Append(p).Append("c>\n    ");
            string cellSpace = row % 5 == 0 ? " xml:space=\"preserve\"" : string.Empty;
            sheet.Append('<').Append(p).Append("c r=\"D").Append(row).Append("\" t=\"n\" s=\"0\"").Append(cellSpace).Append("><").Append(p).Append("v>")
                .Append((row % 400 / 4.0).ToString(System.Globalization.CultureInfo.InvariantCulture)).Append("</").Append(p).Append("v></").Append(p).Append("c>\n");
            sheet.Append("  </").Append(p).Append("row>\n");
        }
        sheet.Append("</").Append(p).Append("sheetData></").Append(p).Append("worksheet>");
        return ReplaceParallelIndexWorksheet(sheet.ToString());
    }

    private static byte[] ReplaceParallelIndexWorksheet(string worksheet) {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddSheet("Data").Cell(1, 1).SetValue("Name");
        using var buffer = new MemoryStream();
        byte[] original = document.ToBytes();
        buffer.Write(original, 0, original.Length);
        using (var package = SpreadsheetDocument.Open(buffer, true)) {
            var part = package.WorkbookPart!.SharedStringTablePart ?? package.WorkbookPart.AddNewPart<SharedStringTablePart>();
            part.SharedStringTable = new SharedStringTable(
                Enumerable.Range(0, 8).Select(index => new SharedStringItem(new Text(index == 0 ? "Name" : "Region " + index))));
            part.SharedStringTable.Save();
        }
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true)) {
            ReplaceAsyncCultureEntry(archive, "xl/worksheets/sheet1.xml", worksheet);
            ReplaceAsyncCultureEntry(archive, "xl/styles.xml", "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/></cellXfs></styleSheet>");
        }
        return buffer.ToArray();
    }

    private static byte[] CreateParallelIndexWorkbook(int invalidSharedStringRow, bool implicitLayout = false) {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddSheet("Data").Cell(1, 1).SetValue("Name");
        using var buffer = new MemoryStream();
        byte[] original = document.ToBytes();
        buffer.Write(original, 0, original.Length);
        using (var package = SpreadsheetDocument.Open(buffer, true)) {
            var part = package.WorkbookPart!.SharedStringTablePart ?? package.WorkbookPart.AddNewPart<SharedStringTablePart>();
            part.SharedStringTable = new SharedStringTable(
                Enumerable.Range(0, 8).Select(index => new SharedStringItem(new Text(index == 0 ? "Name" : "Region " + index))));
            part.SharedStringTable.Save();
        }

        var sheet = new StringBuilder(ParallelIndexRows * 300);
        sheet.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        if (!implicitLayout) sheet.Append("<dimension ref=\"A1:G").Append(ParallelIndexRows).Append("\"/>");
        sheet.Append("<sheetData>")
            .Append("<row r=\"1\" spans=\"1:7\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"inlineStr\"><is><t>Date</t></is></c>")
            .Append("<c r=\"C1\" t=\"inlineStr\"><is><t>Id</t></is></c><c r=\"D1\" t=\"inlineStr\"><is><t>Price</t></is></c>")
            .Append("<c r=\"E1\" t=\"inlineStr\"><is><t>Cost</t></is></c><c r=\"F1\" t=\"inlineStr\"><is><t>Delta</t></is></c>")
            .Append("<c r=\"G1\" t=\"inlineStr\"><is><t>Note</t></is></c></row>");
        for (int row = 2; row <= ParallelIndexRows; row++) {
            // Irregular rows make the background chunk stop and the sequential indexer resume.
            string spans = row % 5003 == 0 ? "1:6" : "1:7";
            // Explicit row numbers would let a misnumbered chunk fall back to the general reader.
            bool implicitRow = implicitLayout;
            if (implicitRow) sheet.Append("<row>");
            else sheet.Append("<row r=\"").Append(row).Append("\" spans=\"").Append(spans).Append("\">");
            string Cell(char column) => implicitRow ? "<c" : "<c r=\"" + column + row + "\"";
            int sharedString = row == invalidSharedStringRow ? 99 : 1 + row % 7;
            sheet.Append(Cell('A')).Append(" t=\"s\"><v>").Append(sharedString).Append("</v></c>");
            sheet.Append(Cell('B')).Append(" s=\"1\"><v>").Append(40000 + row % 3000).Append(row % 4 == 0 ? ".5" : "").Append("</v></c>");
            if (implicitRow && row % 7919 == 0) sheet.Append("<c/>");
            else if (row % 7919 != 0) sheet.Append(Cell('C')).Append("><v>").Append(400_000_000 + row * 7).Append("</v></c>");
            sheet.Append(Cell('D')).Append("><v>").Append(row % 3 == 0 ? "35.840000000000003" : (row % 1000 / 8.0).ToString(System.Globalization.CultureInfo.InvariantCulture)).Append("</v></c>");
            if (row == 17_011) sheet.Append(Cell('E')).Append("><f>1+1</f><v>2</v></c>");
            else sheet.Append(Cell('E')).Append("><v>165258.23999999999</v></c>");
            sheet.Append(Cell('F')).Append("><v>").Append(row % 2 == 0 ? "-0.25" : "-0").Append("</v></c>");
            if (row == 13_001) sheet.Append(Cell('G')).Append(" t=\"inlineStr\"><is><t>inline</t></is></c>");
            else sheet.Append(Cell('G')).Append(" t=\"s\"><v>").Append(row % 5).Append("</v></c>");
            sheet.Append("</row>");
            if (row % 9001 == 0) sheet.Append('\n');
        }
        sheet.Append("</sheetData></worksheet>");

        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true)) {
            ReplaceAsyncCultureEntry(archive, "xl/worksheets/sheet1.xml", sheet.ToString());
            ReplaceAsyncCultureEntry(archive, "xl/styles.xml", "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/></cellXfs></styleSheet>");
        }
        return buffer.ToArray();
    }
}
