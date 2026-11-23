using System.Globalization;
using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    // Large enough (about 19 MB of worksheet XML) that the reader holds only a window of the inflated part.
    private const int LargeWorksheetRows = 110_000;

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_LargeWorksheetReturnsEveryValue(bool dimension) {
        string path = CreateWindowedIndexWorkbook(LargeWindowWorksheet(dimension, lastRowCells: null, tail: null));
        try {
            using var reader = ExcelDocument.OpenReader(path);
            Assert.Equal(4, reader.FieldCount);
            Assert.Equal(new[] { "Name", "Id", "Date", "Value" }, Enumerable.Range(0, 4).Select(reader.GetName));
            int count = 0;
            while (reader.Read()) {
                int row = count + 2;
                Assert.Equal("name " + row.ToString(CultureInfo.InvariantCulture) + " & co", reader.GetString(0));
                Assert.Equal(row, reader.GetInt32(1));
                if (row % 1000 == 0) {
                    Assert.True(reader.IsDBNull(2));
                } else {
                    Assert.Equal(DateTime.FromOADate(45292.25), reader.GetDateTime(2));
                }
                Assert.Equal(row * 2.0, reader.GetDouble(3));
                count++;
            }
            Assert.Equal(LargeWorksheetRows, count);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_LargeWorksheetRejectsLateInvalidInputBeforeTheFirstRow(bool malformedXml) {
        string path = CreateWindowedIndexWorkbook(malformedXml
            ? LargeWindowWorksheet(dimension: false, lastRowCells: null, tail: "<broken>")
            : LargeWindowWorksheet(dimension: false, lastRowCells: "<c s=\"9\"/>", tail: null));
        try {
            if (malformedXml) {
                Assert.Throws<XmlException>(() => { using var reader = ExcelDocument.OpenReader(path); _ = reader.FieldCount; });
            } else {
                Assert.Throws<InvalidDataException>(() => { using var reader = ExcelDocument.OpenReader(path); _ = reader.FieldCount; });
            }
        } finally {
            File.Delete(path);
        }
    }

    private static string LargeWindowWorksheet(bool dimension, string? lastRowCells, string? tail) {
        int lastRow = LargeWorksheetRows + 1;
        var xml = new StringBuilder(lastRow * 180);
        xml.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        if (dimension) xml.Append("<dimension ref=\"A1:D").Append(lastRow).Append("\"/>");
        xml.Append("<sheetData>").Append(WindowedIndexHeader.Replace("<row>", "<row r=\"1\">"));
        for (int row = 2; row <= lastRow; row++) {
            string r = row.ToString(CultureInfo.InvariantCulture);
            xml.Append("<row r=\"").Append(r).Append("\"><c r=\"A").Append(r).Append("\" t=\"inlineStr\"><is><t>name ")
                .Append(r).Append(" &amp; co</t></is></c><c r=\"B").Append(r).Append("\"><v>").Append(r).Append("</v></c>");
            if (row % 1000 != 0) xml.Append("<c r=\"C").Append(r).Append("\" s=\"1\"><v>45292.25</v></c>");
            xml.Append("<c r=\"D").Append(r).Append("\"><f>B").Append(r).Append("*2</f><v>")
                .Append((row * 2).ToString(CultureInfo.InvariantCulture)).Append("</v></c>");
            if (row == lastRow && lastRowCells != null) xml.Append(lastRowCells);
            xml.Append("</row>");
        }
        xml.Append("</sheetData>").Append(tail).Append("</worksheet>");
        return xml.ToString();
    }
}
