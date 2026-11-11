using System.Data.Common;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    private const int XlsbSinglePassRows = 30_000;

    /// <summary>Single-pass XLSB reading returns the default reader's schema and values.</summary>
    [Fact]
    public void XlsbSinglePassReadingMatchesTheDefaultReader() {
        byte[] workbook = CreateSinglePassXlsb(extraColumnRow: 0);
        using DbDataReader expected = ExcelDocument.OpenReader(workbook, new ExcelReaderOptions { NumericAsDecimal = true });
        using DbDataReader actual = ExcelDocument.OpenReader(workbook,
            new ExcelReaderOptions { NumericAsDecimal = true, EnableSinglePassReading = true });
        int rows = 0;
        while (expected.Read()) {
            Assert.True(actual.Read());
            Assert.Equal(expected.FieldCount, actual.FieldCount);
            for (int ordinal = 0; ordinal < expected.FieldCount; ordinal++) {
                Assert.Equal(expected.GetValue(ordinal), actual.GetValue(ordinal));
            }
            rows++;
        }
        Assert.False(actual.Read());
        Assert.Equal(XlsbSinglePassRows - 1, rows);
        Assert.Equal(Enumerable.Range(0, expected.FieldCount).Select(expected.GetName),
            Enumerable.Range(0, actual.FieldCount).Select(actual.GetName));
    }

    /// <summary>
    /// Single-pass XLSB reading returns exactly the valid rows before a malformed cell, then raises the
    /// error the default reader raises before its first row.
    /// </summary>
    [Fact]
    public void XlsbSinglePassReadingReturnsTheValidPrefixThenTheSameError() {
        int invalidRow = XlsbSinglePassRows - 7;
        byte[] workbook = BreakSharedStringIndex(CreateSinglePassXlsb(extraColumnRow: 0), invalidRow);
        InvalidDataException expected = Assert.Throws<InvalidDataException>(() => {
            using DbDataReader reader = ExcelDocument.OpenReader(workbook, new ExcelReaderOptions { NumericAsDecimal = true });
            reader.Read();
        });

        using DbDataReader reference = ExcelDocument.OpenReader(CreateSinglePassXlsb(extraColumnRow: 0),
            new ExcelReaderOptions { NumericAsDecimal = true });
        using DbDataReader singlePass = ExcelDocument.OpenReader(workbook,
            new ExcelReaderOptions { NumericAsDecimal = true, EnableSinglePassReading = true });
        int delivered = 0;
        InvalidDataException actual = Assert.Throws<InvalidDataException>(() => {
            while (singlePass.Read()) {
                Assert.True(reference.Read());
                for (int ordinal = 0; ordinal < reference.FieldCount; ordinal++) {
                    Assert.Equal(reference.GetValue(ordinal), singlePass.GetValue(ordinal));
                }
                delivered++;
            }
        });
        Assert.Equal(expected.Message, actual.Message);
#if NET8_0_OR_GREATER
        // The header is row 1, so every data row before the malformed one was returned.
        Assert.Equal(invalidRow - 2, delivered);
#else
        Assert.Equal(0, delivered);
#endif
    }

    /// <summary>
    /// A later cell outside the header's columns widens the default reader's table, so single-pass
    /// reading stops at that row rather than return rows with a different shape.
    /// </summary>
    [Fact]
    public void XlsbSinglePassReadingStopsAtACellOutsideTheHeaderColumns() {
        int extraRow = XlsbSinglePassRows - 11;
        byte[] workbook = CreateSinglePassXlsb(extraColumnRow: extraRow);
        using (DbDataReader reader = ExcelDocument.OpenReader(workbook, new ExcelReaderOptions { NumericAsDecimal = true })) {
            Assert.Equal(8, reader.FieldCount);
        }

        using DbDataReader singlePass = ExcelDocument.OpenReader(workbook,
            new ExcelReaderOptions { NumericAsDecimal = true, EnableSinglePassReading = true });
#if NET8_0_OR_GREATER
        int delivered = 0;
        Assert.Throws<NotSupportedException>(() => {
            while (singlePass.Read()) delivered++;
        });
        Assert.Equal(extraRow - 2, delivered);
#else
        Assert.Equal(8, singlePass.FieldCount);
        int rows = 0;
        while (singlePass.Read()) rows++;
        Assert.Equal(XlsbSinglePassRows - 1, rows);
#endif
    }

    /// <summary>Disposing a single-pass XLSB reader early stops and joins its background work.</summary>
    [Fact]
    public void XlsbSinglePassReaderDisposesWhileTheWorksheetIsStillBeingRead() {
        byte[] workbook = CreateSinglePassXlsb(extraColumnRow: 0);
        for (int attempt = 0; attempt < 5; attempt++) {
            using DbDataReader reader = ExcelDocument.OpenReader(workbook,
                new ExcelReaderOptions { EnableSinglePassReading = true });
            for (int row = 0; row < attempt * 7; row++) Assert.True(reader.Read());
        }
    }

    /// <summary>
    /// Moving to the next worksheet while the current one is still inflating reads the next sheet as the
    /// default reader does, so the two sheets never share the package archive at the same time.
    /// </summary>
    [Fact]
    public void XlsbSinglePassNextResultReadsTheNextSheetWhileTheCurrentOneIsStillInflating() {
        byte[] workbook = CreateSinglePassXlsb(extraColumnRow: 0, sheetCount: 2);
        var expected = new List<object>();
        using (DbDataReader second = ExcelDocument.OpenReader(workbook, new ExcelReaderOptions { SheetIndex = 1 })) {
            while (second.Read()) {
                for (int ordinal = 0; ordinal < second.FieldCount; ordinal++) expected.Add(second.GetValue(ordinal));
            }
        }

        for (int attempt = 0; attempt < 5; attempt++) {
            using DbDataReader reader = ExcelDocument.OpenReader(workbook, new ExcelReaderOptions { EnableSinglePassReading = true });
            for (int row = 0; row < attempt * 3; row++) Assert.True(reader.Read());
            Assert.True(reader.NextResult());
            var actual = new List<object>();
            while (reader.Read()) {
                for (int ordinal = 0; ordinal < reader.FieldCount; ordinal++) actual.Add(reader.GetValue(ordinal));
            }
            Assert.Equal(expected, actual);
        }
    }

    private static byte[] CreateSinglePassXlsb(int extraColumnRow, int sheetCount = 1) {
        using ExcelDocument document = ExcelDocument.Create();
        for (int index = 0; index < sheetCount; index++) {
            FillSinglePassSheet(document.AddSheet(index == 0 ? "Data" : "Data" + (index + 1)), extraColumnRow);
        }
        return document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb, XlsbUseSharedStrings = true });
    }

    private static void FillSinglePassSheet(ExcelSheet sheet, int extraColumnRow) {
        string[] headers = { "Region", "Date", "Id", "Price", "Cost", "Note" };
        for (int column = 0; column < headers.Length; column++) sheet.Cell(1, column + 1).SetValue(headers[column]);
        for (int row = 2; row <= XlsbSinglePassRows; row++) {
            // A gap row the reader returns as blank. The writer fills it with blank cells across the used
            // columns, so the late-column workbook leaves it out to keep that cell the only one outside.
            if (row % 4099 == 0 && extraColumnRow == 0) continue;
            sheet.Cell(row, 1).SetValue("Region " + row % 7);
            sheet.Cell(row, 2).SetValue(new DateTime(2020, 1, 1).AddDays(row % 900));
            if (row % 97 != 0) sheet.Cell(row, 3).SetValue(row);
            sheet.Cell(row, 4).SetValue(row % 400 / 4.0m);
            sheet.Cell(row, 5).SetValue(row * 1.25 / 3);
            sheet.Cell(row, 6).SetValue("note " + row % 11);
            if (row == extraColumnRow) sheet.Cell(row, 8).SetValue("late");
        }
    }

    /// <summary>Points the first shared-string cell of a 1-based worksheet row at a missing item.</summary>
    private static byte[] BreakSharedStringIndex(byte[] workbook, int row) {
        using var buffer = new MemoryStream();
        buffer.Write(workbook, 0, workbook.Length);
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.Entries.Single(candidate =>
                candidate.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal) && candidate.FullName.EndsWith(".bin", StringComparison.Ordinal));
            byte[] sheet;
            using (var input = new MemoryStream()) {
                using (Stream stream = entry.Open()) stream.CopyTo(input);
                sheet = input.ToArray();
            }
            int position = 0;
            bool inRow = false;
            bool patched = false;
            while (position < sheet.Length && !patched) {
                int type = sheet[position++];
                if ((type & 0x80) != 0) type = (type & 0x7F) | (sheet[position++] & 0x7F) << 7;
                int size = 0;
                for (int shift = 0; ; shift += 7) {
                    int value = sheet[position++];
                    size |= (value & 0x7F) << shift;
                    if ((value & 0x80) == 0) break;
                }
                if (type == 0) inRow = BitConverter.ToInt32(sheet, position) == row - 1; // BrtRowHdr
                else if (type == 7 && inRow) { // BrtCellIsst: column, style, then the item index
                    BitConverter.GetBytes(int.MaxValue).CopyTo(sheet, position + 8);
                    patched = true;
                }
                position += size;
            }
            Assert.True(patched);
            entry.Delete();
            using Stream output = archive.CreateEntry(entry.FullName).Open();
            output.Write(sheet, 0, sheet.Length);
        }
        return buffer.ToArray();
    }
}
