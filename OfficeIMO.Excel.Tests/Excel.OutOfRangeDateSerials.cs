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
    /// reads and conversions return the serial number for that cell and keep reading the rest of the sheet.
    /// </summary>
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    [InlineData(".xls")]
    public void DateFormattedSerialsWithoutADate_ReadAsTheirNumber(string extension) {
        string source = CreateOutOfRangeDateWorkbook();
        string path = source;
        try {
            if (extension != ".xlsx") {
                path = Path.ChangeExtension(source, extension);
                using var converted = ExcelDocument.Load(source);
                converted.Save(path);
            }

            using (DbDataReader reader = ExcelDocument.OpenReader(path)) {
                Assert.True(reader.Read());
                Assert.Equal(3_000_000D, Assert.IsType<double>(reader.GetValue(0)));
                Assert.Equal(1D, Convert.ToDouble(reader.GetValue(1)));
                Assert.True(reader.Read());
                Assert.Equal(new DateTime(2024, 1, 1), reader.GetValue(0));
                Assert.False(reader.Read());
            }

            // A sampled dateless serial infers the type that GetValue returns for it.
            using (DbDataReader inferred = ExcelDocument.OpenReader(path, new ExcelReaderOptions { InferSchema = true, SchemaSampleRows = 1 })) {
                Assert.Equal(typeof(double), inferred.GetFieldType(0));
                Assert.True(inferred.Read());
                Assert.Equal(3_000_000D, Assert.IsType<double>(inferred.GetValue(0)));
            }

            if (extension == ".xlsx") {
                using var workbook = ExcelDocumentReader.Open(path);
                object?[] row = workbook.GetSheet(1).ReadRows("A2:B2").Single()!;
                Assert.Equal(3_000_000D, row[0]);

                using var document = ExcelDocument.Load(path);
                Assert.Contains("3000000", document.Sheets[0].ToCsv().ToString(), StringComparison.Ordinal);
            }
        } finally {
            File.Delete(source);
            if (!ReferenceEquals(path, source)) File.Delete(path);
        }
    }

    private static string CreateOutOfRangeDateWorkbook() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.OutOfRangeDate.{Guid.NewGuid():N}.xlsx");
        using (var document = ExcelDocument.Create(path)) {
            document.AddSheet("Data").Cell(1, 1).SetValue("Date");
            document.Save();
        }

        const string worksheet = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
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
        foreach ((string name, string content) in new[] { ("xl/worksheets/sheet1.xml", worksheet), ("xl/styles.xml", styles) }) {
            archive.GetEntry(name)?.Delete();
            using Stream entry = archive.CreateEntry(name).Open();
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            entry.Write(bytes, 0, bytes.Length);
        }

        return path;
    }
}
