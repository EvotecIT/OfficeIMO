using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_RepeatedImplicitTagsRetainChangedTypesStylesAndMissingValues(bool treatDates) {
        string[] rows = {
            "<c t=\"inlineStr\"><is><t>alpha</t></is></c><c t=\"b\"><v>1</v></c><c s=\"1\"><v>45292</v></c><c/><c><v>101</v></c>",
            "<c t=\"inlineStr\"><is><t>beta</t></is></c><c t=\"b\"><v>0</v></c><c s=\"1\"><v>45293</v></c><c/><c><v>102</v></c>",
            "<c t=\"str\"><v>typed</v></c><c t=\"n\"><v>7</v></c><c s=\"0\"><v>9</v></c><c t=\"inlineStr\"><is><t></t></is></c><c><v>103</v></c>",
            "<c t=\"inlineStr\"><is><t>AT&amp;T</t></is></c><c t=\"n\"><v>8</v></c><c s=\"0\"><v>10</v></c><c/>",
            "<c t=\"inlineStr\"><is><t>end</t></is></c><c t=\"b\"><v>1</v></c><c s=\"1\"><v>45294</v></c><c/><c><v>104</v></c>",
        };
        string path = CreateImplicitCellTagWorkbook(rows, "Name", "Kind", "Date", "Empty", "Tail");
        try {
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes("""
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs>
                </styleSheet>
                """));
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { TreatDatesUsingNumberFormat = treatDates });
            string[] names = { "alpha", "beta", "typed", "AT&T", "end" };
            object[] kinds = { true, false, 7d, 8d, true };
            double[] dates = { 45292, 45293, 9, 10, 45294 };
            for (int row = 0; row < rows.Length; row++) {
                Assert.True(reader.Read());
                Assert.Equal(names[row], reader.GetString(0));
                Assert.Equal(kinds[row], reader.GetValue(1));
                bool dateStyle = row is 0 or 1 or 4;
                object expectedDate = treatDates && dateStyle ? (object)DateTime.FromOADate(dates[row]) : dates[row];
                Assert.Equal(expectedDate, reader.GetValue(2));
                if (row == 2) Assert.Equal(string.Empty, reader.GetString(3));
                else Assert.True(reader.IsDBNull(3));
                if (row == 3) Assert.True(reader.IsDBNull(4));
                else Assert.Equal(row == 4 ? 104 : 101 + row, reader.GetInt32(4));
            }
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RepeatedImplicitTagsRejectAChangedMissingStyleBeforeDelivery() {
        string path = CreateImplicitCellTagWorkbook(new[] {
            "<c s=\"0\"><v>1</v></c>",
            "<c s=\"0\"><v>2</v></c>",
            "<c s=\"9\"><v>3</v></c>",
        }, "Value");
        try {
            var exception = Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(path));
            Assert.Contains("cell A4 references a missing cell style", exception.Message);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RepeatedImplicitSharedStringTagsValidateEachIndexBeforeDelivery() {
        string path = CreateImplicitCellTagWorkbook(new[] {
            "<c t=\"s\"><v>0</v></c>",
            "<c t=\"s\"><v>1</v></c>",
            "<c t=\"s\"><v>2</v></c>",
        }, "Value");
        try {
            using (var package = SpreadsheetDocument.Open(path, true)) {
                var part = package.WorkbookPart!.SharedStringTablePart
                    ?? package.WorkbookPart.AddNewPart<SharedStringTablePart>();
                part.SharedStringTable = new SharedStringTable(
                    new SharedStringItem(new Text("first")), new SharedStringItem(new Text("second")));
                part.SharedStringTable.Save();
            }
            var exception = Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(path));
            Assert.Contains("cell A4 references a missing shared string", exception.Message);
        } finally {
            File.Delete(path);
        }
    }

    private static string CreateImplicitCellTagWorkbook(string[] rows, params string[] headers) {
        string path = CreateCompactFastPathWorkbook();
        string header = string.Concat(headers.Select(value => $"<c t=\"inlineStr\"><is><t>{value}</t></is></c>"));
        string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + $"<dimension ref=\"A1:{(char)('A' + headers.Length - 1)}{rows.Length + 1}\"/><sheetData><row>"
            + header + "</row>" + string.Concat(rows.Select(row => "<row>" + row + "</row>"))
            + "</sheetData></worksheet>";
        ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
        return path;
    }
}
