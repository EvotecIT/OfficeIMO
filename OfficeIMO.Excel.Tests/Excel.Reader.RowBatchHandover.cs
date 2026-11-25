using System.Data.Common;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    private const int HandoverRows = 30_000;

    /// <summary>
    /// A worksheet that turns to markup the compact indexer does not handle, early or after many rows were returned,
    /// keeps reading through the general XML reader with the same values as a reader that validated it first.
    /// </summary>
    [Theory]
    [InlineData(2, "<![CDATA[early & cdata]]>", "early & cdata")]
    [InlineData(HandoverRows - 5_000, "<![CDATA[late & cdata]]>", "late & cdata")]
    [InlineData(HandoverRows - 5_000, "  indented <!-- comment -->", "  indented ")]
    public void OpenReader_MarkupTheFastReaderDeclinesKeepsTheSameValues(int markupRow, string markupText, string expectedText) {
        string path = CreateHandoverWorkbook(markupRow, markupText, dimensionColumns: "B");
        try {
            AssertReadsLikeValidatingFirst(path, markupRow, expectedText);
        } finally {
            File.Delete(path);
        }
    }

    /// <summary>More trailing content than the read window holds after the last row still ends the rows normally.</summary>
    [Fact]
    public void OpenReader_LargeTrailingWhitespaceAfterRowsKeepsTheSameValues() {
        string path = CreateHandoverWorkbook(markupRow: 0, markupText: string.Empty, dimensionColumns: "B",
            trailingWhitespaceBytes: 65 * 1024 * 1024);
        try {
            AssertReadsLikeValidatingFirst(path, markupRow: 0, expectedText: string.Empty);
        } finally {
            File.Delete(path);
        }
    }

    /// <summary>
    /// A declared dimension narrower than a later row is wrong: the rows before it are returned and the read that
    /// reaches it fails, while validating first reads every cell of the actual range.
    /// </summary>
    [Fact]
    public void OpenReader_CellOutsideTheDeclaredDimensionFailsAtThatRow() {
        int wideRow = HandoverRows - 100;
        string path = CreateHandoverWorkbook(markupRow: 0, markupText: string.Empty, dimensionColumns: "B", extraCellRow: wideRow);
        try {
            using (DbDataReader reader = ExcelDocument.OpenReader(path)) {
                Exception? failure = ReadUntilError(reader, out int rows);
                Assert.Contains(nameof(ExcelReaderOptions.ValidateWorksheetBeforeFirstRow),
                    Assert.IsType<InvalidDataException>(failure).Message);
                Assert.True(rows > 0 && rows < wideRow - 1);
            }

            using DbDataReader validated = ExcelDocument.OpenReader(path, new ExcelReaderOptions { ValidateWorksheetBeforeFirstRow = true });
            Assert.Equal(3, validated.FieldCount);
            int validatedRows = 0;
            while (validated.Read()) validatedRows++;
            Assert.Equal(HandoverRows - 1, validatedRows);
        } finally {
            File.Delete(path);
        }
    }

    private static void AssertReadsLikeValidatingFirst(string path, int markupRow, string expectedText) {
        using DbDataReader expected = ExcelDocument.OpenReader(path, new ExcelReaderOptions { ValidateWorksheetBeforeFirstRow = true });
        using DbDataReader actual = ExcelDocument.OpenReader(path);
        Assert.Equal(expected.FieldCount, actual.FieldCount);
        int rows = 0;
        while (expected.Read()) {
            Assert.True(actual.Read());
            rows++;
            for (int ordinal = 0; ordinal < expected.FieldCount; ordinal++) {
                Assert.Equal(expected.GetValue(ordinal), actual.GetValue(ordinal));
            }
            if (rows + 1 == markupRow) Assert.Equal(expectedText, actual.GetString(0));
        }
        Assert.False(actual.Read());
        Assert.Equal(HandoverRows - 1, rows);
    }

    private static string CreateHandoverWorkbook(int markupRow, string markupText, string dimensionColumns,
        int trailingWhitespaceBytes = 0, int extraCellRow = 0) {
        string path = CreateCompactFastPathWorkbook();
        var xml = new StringBuilder(HandoverRows * 120);
        xml.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:")
            .Append(dimensionColumns).Append(HandoverRows.ToString(CultureInfo.InvariantCulture)).Append("\"/><sheetData>");
        for (int row = 1; row <= HandoverRows; row++) {
            string r = row.ToString(CultureInfo.InvariantCulture);
            string text = row == 1 ? "Name" : row == markupRow ? markupText : "Item " + r;
            xml.Append("<row r=\"").Append(r).Append("\"><c r=\"A").Append(r).Append("\" t=\"inlineStr\"><is><t>")
                .Append(text).Append("</t></is></c><c r=\"B").Append(r).Append("\"><v>")
                .Append(row == 1 ? "0" : (row * 3).ToString(CultureInfo.InvariantCulture)).Append("</v></c>");
            if (row == extraCellRow) xml.Append("<c r=\"C").Append(r).Append("\"><v>7</v></c>");
            xml.Append("</row>");
        }
        xml.Append("</sheetData>");
        byte[] head = Encoding.UTF8.GetBytes(xml.ToString());
        byte[] tail = Encoding.UTF8.GetBytes("</worksheet>");
        using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update)) {
            archive.GetEntry("xl/worksheets/sheet1.xml")!.Delete();
            using Stream destination = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Optimal).Open();
            destination.Write(head, 0, head.Length);
            byte[] padding = Enumerable.Repeat((byte)' ', 64 * 1024).ToArray();
            for (int written = 0; written < trailingWhitespaceBytes; written += padding.Length) destination.Write(padding, 0, padding.Length);
            destination.Write(tail, 0, tail.Length);
        }
        return path;
    }
}
