using System.Data.Common;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    /// <summary>
    /// A date-formatted serial past 9999-12-31 has no <see cref="DateTime"/>; Excel displays it as ####. Untyped
    /// reads and conversions return the serial as a <see cref="double"/> in every reader and format, whatever
    /// <see cref="ExcelReaderOptions.NumericAsDecimal"/> says, and keep reading the rest of the sheet. The UTF-16
    /// worksheet goes through the XML reader instead of the UTF-8 one.
    /// </summary>
    [Theory]
    [InlineData(".xlsx", false, false)]
    [InlineData(".xlsx", false, true)]
    [InlineData(".xlsx", true, false)]
    [InlineData(".xlsx", true, true)]
    [InlineData(".xlsb", false, false)]
    [InlineData(".xlsb", false, true)]
    [InlineData(".xls", false, false)]
    [InlineData(".xls", false, true)]
    public void DateFormattedSerialsWithoutADate_ReadAsTheirNumber(string extension, bool utf16, bool numericAsDecimal) {
        string source = CreateOutOfRangeDateWorkbook(utf16);
        string path = source;
        try {
            if (extension != ".xlsx") {
                path = Path.ChangeExtension(source, extension);
                using var converted = ExcelDocument.Load(source);
                converted.Save(path);
            }

            using (DbDataReader reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions { NumericAsDecimal = numericAsDecimal })) {
                Assert.True(reader.Read());
                Assert.Equal(3_000_000D, Assert.IsType<double>(reader.GetValue(0)));
                Assert.Equal("3000000", reader.GetString(0));
                Assert.Throws<InvalidCastException>(() => reader.GetDateTime(0));
                Assert.Equal(3_000_000D, reader.GetDouble(0));
                Assert.True(reader.Read());
                Assert.Equal(new DateTime(2024, 1, 1), reader.GetValue(0));
                Assert.False(reader.Read());
            }

            // A sampled dateless serial infers the type that GetValue returns for it.
            using (DbDataReader inferred = ExcelDocument.OpenReader(path, new ExcelReaderOptions {
                InferSchema = true, SchemaSampleRows = 1, NumericAsDecimal = numericAsDecimal })) {
                Assert.Equal(typeof(double), inferred.GetFieldType(0));
                Assert.True(inferred.Read());
                Assert.Equal(3_000_000D, Assert.IsType<double>(inferred.GetValue(0)));
            }

            if (extension == ".xlsx") {
                using var workbook = ExcelDocumentReader.Open(path, new ExcelReaderOptions { NumericAsDecimal = numericAsDecimal });
                object?[] row = workbook.GetSheet(1).ReadRows("A2:B2").Single()!;
                Assert.Equal(3_000_000D, Assert.IsType<double>(row[0]));
                Dictionary<string, object?> record = workbook.GetSheet(1).ReadObjects("A1:B2").Single();
                Assert.Equal(3_000_000D, Assert.IsType<double>(record["Date"]));

                using var document = ExcelDocument.Load(path);
                Assert.Contains("3000000", document.Sheets[0].ToCsv().ToString(), StringComparison.Ordinal);
            }
        } finally {
            File.Delete(source);
            if (!ReferenceEquals(path, source)) File.Delete(path);
        }
    }

    private static string CreateOutOfRangeDateWorkbook(bool utf16 = false) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.OutOfRangeDate.{Guid.NewGuid():N}.xlsx");
        using (var document = ExcelDocument.Create(path)) {
            document.AddSheet("Data").Cell(1, 1).SetValue("Date");
            document.Save();
        }

        string worksheet = $"<?xml version=\"1.0\" encoding=\"{(utf16 ? "utf-16" : "utf-8")}\"?>" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:B3\"/><sheetData>" +
            "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Date</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Amount</t></is></c></row>" +
            "<row r=\"2\"><c r=\"A2\" s=\"1\"><v>3000000</v></c><c r=\"B2\"><v>1</v></c></row>" +
            "<row r=\"3\"><c r=\"A3\" s=\"1\"><v>45292</v></c><c r=\"B3\"><v>2</v></c></row>" +
            "</sheetData></worksheet>";
        const string styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>" +
            "<borders count=\"1\"><border/></borders><cellStyleXfs count=\"1\"><xf/></cellStyleXfs>" +
            "<cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\" applyNumberFormat=\"1\"/></cellXfs></styleSheet>";
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        foreach ((string name, string content, bool wide) in new[] { ("xl/worksheets/sheet1.xml", worksheet, utf16), ("xl/styles.xml", styles, false) }) {
            archive.GetEntry(name)?.Delete();
            using Stream entry = archive.CreateEntry(name).Open();
            byte[] bytes = wide ? Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(content)).ToArray() : Encoding.UTF8.GetBytes(content);
            entry.Write(bytes, 0, bytes.Length);
        }

        return path;
    }
}
