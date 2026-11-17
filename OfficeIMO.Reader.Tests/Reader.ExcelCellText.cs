using System.Globalization;
using OfficeIMO.Excel;
using OfficeIMO.Reader.Excel;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed partial class ReaderExcelCellTextTests {
    private static readonly string[] Headers = {
        "Share", "Amount", "Date", "Time", "Fraction", "Scientific", "Code", "General", "Active"
    };

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xls")]
    [InlineData(".xlsb")]
    public void FormattedCellTextFollowsNumberFormatsRegardlessOfTheMachineCulture(string extension) {
        byte[] workbook = BuildWorkbook(extension);
        CultureInfo original = CultureInfo.CurrentCulture;
        IReadOnlyList<string> row;
        try {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            row = ReadFirstDataRow(workbook, extension, new ReaderExcelOptions());
        } finally {
            CultureInfo.CurrentCulture = original;
        }

        Assert.Equal(new[] { "3.2%", "$31,200", "2024-01-15", "1:30 PM", "1 1/4", "1.23E+04", "00123", "0.3", "TRUE" }, row);
    }

    [Fact]
    public void LargeWorksheetsStreamTheSameFormattedText() {
        // Past 4,096 rows the reader streams the worksheet XML directly instead of buffering parsed rows.
        const int dataRows = 5_000;
        byte[] workbook = BuildWorkbook(".xlsx", dataRows);

        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddExcelHandler().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(workbook, "large.xlsx",
            new ReaderOptions { MaxTableRows = dataRows, MaxChars = 1_000_000 });

        IReadOnlyList<string>[] rows = result.Chunks.SelectMany(chunk => chunk.Tables!).SelectMany(table => table.Rows).ToArray();
        Assert.Equal(dataRows, rows.Length);
        Assert.All(new[] { rows[0], rows[dataRows - 1] }, row =>
            Assert.Equal(new[] { "3.2%", "$31,200", "2024-01-15", "1:30 PM", "1 1/4", "1.23E+04", "00123", "0.3", "TRUE" }, row));
    }

    [Fact]
    public void WorkbookTablesUseTheSameFormattedText() {
        byte[] workbook = BuildWorkbook(".xlsx", withTable: true);

        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddExcelHandler().Build()
            .ReadDocument(workbook, "table.xlsx");

        ReaderTable table = Assert.Single(result.Tables, candidate => candidate.Kind == "excel-table");
        Assert.Equal(new[] { "3.2%", "$31,200", "2024-01-15", "1:30 PM", "1 1/4", "1.23E+04", "00123", "0.3", "TRUE" }, Assert.Single(table.Rows));
    }

    [Fact]
    public void RawCellTextKeepsStoredValues() {
        byte[] workbook = BuildWorkbook(".xlsx");

        IReadOnlyList<string> row = ReadFirstDataRow(workbook, ".xlsx", new ReaderExcelOptions { CellText = ReaderExcelCellText.Raw });

        // Raw numbers use the runtime's default invariant text, which is shortest round-trip on .NET and 15 digits on .NET Framework.
        Assert.Equal(new[] {
            "0.032", "31200", "2024-01-15T00:00:00.0000000", "1899-12-31T13:30:00.0000000", "1.25", "12345.678", "00123",
            (0.1d + 0.2d).ToString(CultureInfo.InvariantCulture), "True"
        }, row);
    }

    [Fact]
    public void CellTextCultureLocalizesSeparatorsButKeepsLiteralSymbols() {
        byte[] workbook = BuildWorkbook(".xlsx");

        IReadOnlyList<string> row = ReadFirstDataRow(workbook, ".xlsx",
            new ReaderExcelOptions { CellTextCulture = CultureInfo.GetCultureInfo("de-DE") });

        Assert.Equal("3,2%", row[0]);
        Assert.Equal("$31.200", row[1]);
        Assert.Equal("2024-01-15", row[2]);
        Assert.Equal("1,23E+04", row[5]);
        Assert.Equal("0,3", row[7]);
    }

    private static IReadOnlyList<string> ReadFirstDataRow(byte[] workbook, string extension, ReaderExcelOptions options) {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddExcelHandler(options).Build();
        OfficeDocumentReadResult result = reader.ReadDocument(workbook, "formats" + extension);
        ReaderTable table = Assert.Single(Assert.Single(result.Chunks).Tables!);
        Assert.Equal(Headers, table.Columns);
        return Assert.Single(table.Rows);
    }

    private static byte[] BuildWorkbook(string extension, int dataRows = 1, bool withTable = false) {
        string path = Path.Combine(Path.GetTempPath(), "reader-cell-text-" + Guid.NewGuid().ToString("N") + extension);
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddSheet("Data");
                for (int column = 0; column < Headers.Length; column++) sheet.Cell(1, column + 1).SetValue(Headers[column]);
                for (int row = 2; row <= dataRows + 1; row++) {
                    sheet.Cell(row, 1).SetValue(0.032d).SetNumberFormat("0.0%");
                    sheet.Cell(row, 2).SetValue(31200d).SetNumberFormat("\"$\"#,##0");
                    sheet.Cell(row, 3).SetValue(new DateTime(2024, 1, 15)).SetNumberFormat("yyyy-mm-dd");
                    sheet.Cell(row, 4).SetValue(0.5625d).SetNumberFormat("h:mm AM/PM");
                    sheet.Cell(row, 5).SetValue(1.25d).SetNumberFormat("# ?/?");
                    sheet.Cell(row, 6).SetValue(12345.678d).SetNumberFormat("0.00E+00");
                    sheet.Cell(row, 7).SetValue("00123").SetNumberFormat("@");
                    sheet.Cell(row, 8).SetValue(0.1d + 0.2d);
                    sheet.Cell(row, 9).SetValue(true);
                }
                if (withTable) sheet.Range(1, 1, dataRows + 1, Headers.Length).CreateTable("Formats");
                document.Save(path);
            }
            return File.ReadAllBytes(path);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
