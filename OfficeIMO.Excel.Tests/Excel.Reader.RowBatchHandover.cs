using System.Data.Common;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
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

    /// <summary>
    /// Rows as Excel writes them, with row attributes in another namespace and formulas with attributes, are checked by
    /// the full XML check while they are read and return the same values.
    /// </summary>
    [Fact]
    public void OpenReader_ExcelShapedFormulaRowsKeepTheSameValues() {
        string path = CreateCompactFastPathWorkbook();
        try {
            var xml = new StringBuilder(HandoverRows * 200);
            xml.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"")
                .Append(" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\"")
                .Append(" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">")
                .Append("<dimension ref=\"A1:B").Append(HandoverRows.ToString(CultureInfo.InvariantCulture)).Append("\"/><sheetData>");
            for (int row = 1; row <= HandoverRows; row++) {
                string r = row.ToString(CultureInfo.InvariantCulture);
                xml.Append("<row r=\"").Append(r).Append("\" spans=\"1:2\" x14ac:dyDescent=\"0.25\"><c r=\"A").Append(r).Append("\"><v>")
                    .Append(r).Append("</v></c><c r=\"B").Append(r).Append("\">");
                if (row > 1) xml.Append("<f ca=\"1\">A").Append(r).Append("*2&amp;\"\"</f>");
                xml.Append("<v>").Append((row * 2).ToString(CultureInfo.InvariantCulture)).Append("</v></c></row>");
            }
            xml.Append("</sheetData></worksheet>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml.ToString()));
            AssertReadsLikeValidatingFirst(path, markupRow: 0, expectedText: string.Empty);
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
            // Where the sheet is checked before the first row (.NET Framework), the default read finds the actual range.
            if (ChecksWhileReading) {
                using DbDataReader reader = ExcelDocument.OpenReader(path);
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

    /// <summary>
    /// After a switch to the general XML reader, a row below the declared dimension still fails at that point instead of
    /// ending the rows early.
    /// </summary>
    [Fact]
    public void OpenReader_RowBelowTheDeclaredDimensionAfterTheSwitchFails() {
        int markupRow = HandoverRows - 5_000;
        string path = CreateHandoverWorkbook(markupRow, "<![CDATA[late & cdata]]>", dimensionColumns: "B", extraRows: 1);
        try {
            using DbDataReader reader = ExcelDocument.OpenReader(path);
            Exception? failure = ReadUntilError(reader, out int rows);
            if (!ChecksWhileReading) {
                // Checked before the first row (.NET Framework): every row of the actual range is returned.
                Assert.Null(failure);
                Assert.Equal(HandoverRows, rows);
                return;
            }
            Assert.Contains(nameof(ExcelReaderOptions.ValidateWorksheetBeforeFirstRow),
                Assert.IsType<InvalidDataException>(failure).Message);
            Assert.True(rows > 0 && rows < markupRow);

            using DbDataReader validated = ExcelDocument.OpenReader(path, new ExcelReaderOptions { ValidateWorksheetBeforeFirstRow = true });
            int validatedRows = 0;
            while (validated.Read()) validatedRows++;
            Assert.Equal(HandoverRows, validatedRows);
        } finally {
            File.Delete(path);
        }
    }

    /// <summary>A character XML does not allow in a late cell fails the read that reaches it; that row is not returned.</summary>
    [Fact]
    public void OpenReader_InvalidCharacterInALateCellIsNotReturned() {
        int badRow = HandoverRows - 5_000;
        string path = CreateHandoverWorkbook(badRow, "bad\u0001text", dimensionColumns: "B");
        try {
            using DbDataReader reader = ExcelDocument.OpenReader(path);
            Assert.IsType<XmlException>(ReadUntilError(reader, out int rows));
            // The header is row 1, so rows 2 to badRow - 1 came back.
            Assert.Equal(ChecksWhileReading ? badRow - 2 : 0, rows);
        } finally {
            File.Delete(path);
        }
    }

    /// <summary>A character XML does not allow in a late row's attribute fails the read that reaches that row.</summary>
    [Fact]
    public void OpenReader_InvalidCharacterInALateRowAttributeIsNotReturned() {
        int badRow = HandoverRows - 5_000;
        string path = CreateHandoverWorkbook(badRow, "Item", dimensionColumns: "B", markupRowAttributes: " spans=\"1:2\u0001\"");
        try {
            using DbDataReader reader = ExcelDocument.OpenReader(path);
            Assert.IsType<XmlException>(ReadUntilError(reader, out int rows));
            Assert.Equal(ChecksWhileReading ? badRow - 2 : 0, rows);
        } finally {
            File.Delete(path);
        }
    }

    /// <summary>
    /// When every row waits for the full XML check, as after a formula with attributes, a late XML error still returns
    /// the rows the check read cleanly before it; the error comes from a read close to it.
    /// </summary>
    [Fact]
    public void OpenReader_LateXmlErrorAfterCheckedRowsKeepsTheRowsBeforeIt() {
        int badRow = HandoverRows - 5_000;
        string path = CreateHandoverWorkbook(badRow, "bad\u0001text", dimensionColumns: "B", attributedFormulaRow: 3);
        try {
            using DbDataReader reader = ExcelDocument.OpenReader(path);
            if (!ChecksWhileReading) {
                Assert.IsType<XmlException>(ReadUntilError(reader, out int none));
                Assert.Equal(0, none);
                return;
            }
            Assert.True(reader.Read());
            // A slow consumer holds back the rows still to be copied while the full check reaches the error.
            Thread.Sleep(500);
            Assert.IsType<XmlException>(ReadUntilError(reader, out int rows));
            // The check reads ahead of what it has parsed by at most 64 KB, about 650 of these rows.
            Assert.InRange(rows + 1, badRow - 2 - 800, badRow - 2);
        } finally {
            File.Delete(path);
        }
    }

    /// <summary>
    /// A read canceled while the worksheet switches to the general XML reader can be retried: the next read returns the
    /// same rows as a reader that validated the worksheet first.
    /// </summary>
    [Fact]
    public async Task OpenReader_CanceledReadDuringTheSwitchCanBeRetried() {
        int markupRow = HandoverRows - 5_000;
        // Trailing content makes the full check of the switch long enough for the cancellation to land inside it.
        string path = CreateHandoverWorkbook(markupRow, "<![CDATA[late & cdata]]>", dimensionColumns: "B",
            trailingWhitespaceBytes: 65 * 1024 * 1024);
        try {
            using DbDataReader expected = ExcelDocument.OpenReader(path, new ExcelReaderOptions { ValidateWorksheetBeforeFirstRow = true });
            using DbDataReader actual = ExcelDocument.OpenReader(path);
            int rows = 0;
            while (expected.Read()) {
                bool read;
                if (Math.Abs(rows + 2 - markupRow) <= 2) {
                    using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(5));
                    try {
                        read = await actual.ReadAsync(cancellation.Token);
                    } catch (OperationCanceledException) {
                        read = await actual.ReadAsync(CancellationToken.None);
                    }
                } else {
                    read = actual.Read();
                }
                Assert.True(read);
                rows++;
                for (int ordinal = 0; ordinal < expected.FieldCount; ordinal++) {
                    Assert.Equal(expected.GetValue(ordinal), actual.GetValue(ordinal));
                }
            }
            Assert.False(actual.Read());
            Assert.Equal(HandoverRows - 1, rows);
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
        int trailingWhitespaceBytes = 0, int extraCellRow = 0, int extraRows = 0, string markupRowAttributes = "",
        int attributedFormulaRow = 0) {
        string path = CreateCompactFastPathWorkbook();
        var xml = new StringBuilder(HandoverRows * 120);
        xml.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:")
            .Append(dimensionColumns).Append(HandoverRows.ToString(CultureInfo.InvariantCulture)).Append("\"/><sheetData>");
        for (int row = 1; row <= HandoverRows + extraRows; row++) {
            string r = row.ToString(CultureInfo.InvariantCulture);
            string text = row == 1 ? "Name" : row == markupRow ? markupText : "Item " + r;
            xml.Append("<row r=\"").Append(r).Append('"').Append(row == markupRow ? markupRowAttributes : string.Empty)
                .Append("><c r=\"A").Append(r).Append("\" t=\"inlineStr\"><is><t>")
                .Append(text).Append("</t></is></c><c r=\"B").Append(r).Append("\">")
                .Append(row == attributedFormulaRow ? "<f ca=\"1\">" + (row * 3).ToString(CultureInfo.InvariantCulture) + "</f>" : string.Empty)
                .Append("<v>").Append(row == 1 ? "0" : (row * 3).ToString(CultureInfo.InvariantCulture)).Append("</v></c>");
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
