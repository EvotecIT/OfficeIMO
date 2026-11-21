using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Reader.Excel;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed partial class ReaderExcelCellTextTests {
    [Theory]
    [InlineData(2)]
    [InlineData(5_000)]
    public void CellsWithoutTheirOwnStyleUseTheRowThenColumnFormat(int dataRows) {
        // Column A loses its cell styles and gets a column style; row 3 gets a custom row style; the fraction
        // values are written with a character reference, which the value must be decoded through.
        byte[] workbook = RewriteWorksheet(BuildWorkbook(".xlsx", dataRows), xml => {
            string percentStyle = Regex.Match(xml, "<c r=\"A2\"[^>]*?\\ss=\"(\\d+)\"").Groups[1].Value;
            string currencyStyle = Regex.Match(xml, "<c r=\"B2\"[^>]*?\\ss=\"(\\d+)\"").Groups[1].Value;
            Assert.NotEmpty(percentStyle);
            Assert.NotEmpty(currencyStyle);
            xml = Regex.Replace(xml, "(<c r=\"A\\d+\"[^>]*?)\\ss=\"\\d+\"", "$1");
            xml = Regex.Replace(xml, "<row r=\"3\"", "<row r=\"3\" customFormat=\"1\" s=\"" + currencyStyle + "\"");
            xml = xml.Replace("<v>1.25</v>", "<v>1&#46;25</v>");
            string column = "<col min=\"1\" max=\"1\" width=\"12\" customWidth=\"1\" style=\"" + percentStyle + "\"/>";
            return xml.Contains("<cols>")
                ? xml.Replace("<cols>", "<cols>" + column)
                : xml.Replace("<sheetData", "<cols>" + column + "</cols><sheetData");
        });

        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddExcelHandler().Build()
            .ReadDocument(workbook, "inherited.xlsx", new ReaderOptions { MaxTableRows = dataRows, MaxChars = 1_000_000 });

        IReadOnlyList<string>[] rows = result.Chunks.SelectMany(chunk => chunk.Tables!).SelectMany(table => table.Rows).ToArray();
        Assert.Equal(dataRows, rows.Length);
        Assert.Equal(new[] { "3.2%", "$31,200", "2024-01-15", "1:30 PM", "1 1/4", "1.23E+04", "00123", "0.3", "TRUE" }, rows[0]);
        Assert.Equal("$0", rows[1][0]);
        Assert.All(rows.Where((_, index) => index != 1), row => Assert.Equal("3.2%", row[0]));
        Assert.All(rows, row => Assert.Equal("1 1/4", row[4]));
    }

    [Fact]
    public void DecimalReadsUseTheSameGeneralTextAndCulture() {
        byte[] workbook = BuildWorkbook(".xlsx");

        IReadOnlyList<string> row = ReadFirstDataRow(workbook, ".xlsx", new ReaderExcelOptions {
            CellTextCulture = CultureInfo.GetCultureInfo("de-DE"),
            ReadOptions = new OfficeIMO.Excel.ExcelReaderOptions { NumericAsDecimal = true }
        });

        Assert.Equal("3,2%", row[0]);
        Assert.Equal("0,3", row[7]);
    }

    private static byte[] RewriteWorksheet(byte[] workbook, Func<string, string> rewrite) {
        using var output = new MemoryStream();
        output.Write(workbook, 0, workbook.Length);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();
            entry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false));
            writer.Write(rewrite(xml));
        }
        return output.ToArray();
    }
}
