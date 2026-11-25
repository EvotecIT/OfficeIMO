using System.Globalization;
using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    // About 3.5 MB of worksheet XML; a 64 KB window makes the reader slide it hundreds of times.
    private const int LargeWorksheetRows = 20_000;

    private static ExcelReaderOptions WorksheetWindowOptions(bool windowed) => new ExcelReaderOptions {
        WindowedWorksheetMinimumBytes = windowed ? 1 : int.MaxValue,
        WorksheetWindowBytes = 64 * 1024,
    };

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_WindowedWorksheetReturnsEveryValue(bool dimension) {
        string path = CreateWindowedIndexWorkbook(LargeWindowWorksheet(dimension, lastRowCells: null, tail: null));
        try {
            using var reader = ExcelDocument.OpenReader(path, WorksheetWindowOptions(windowed: true));
            Assert.Equal(new[] { "Name", "Id", "Date", "Value" }, Enumerable.Range(0, reader.FieldCount).Select(reader.GetName));
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
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void OpenDataReader_WindowedWorksheetReportsLateInvalidInput(bool malformedXml, bool dimension, bool validateFirst) {
        string path = CreateWindowedIndexWorkbook(malformedXml
            ? LargeWindowWorksheet(dimension, lastRowCells: null, tail: "<broken>")
            : LargeWindowWorksheet(dimension, lastRowCells: "<c s=\"9\"/>", tail: null));
        try {
            ExcelReaderOptions options = WorksheetWindowOptions(windowed: true);
            options.ValidateWorksheetBeforeFirstRow = validateFirst;
            Type expected = malformedXml ? typeof(XmlException) : typeof(InvalidDataException);
            if (validateFirst || !dimension) {
                // Without a declared dimension the columns are known only after the whole worksheet, so it is
                // validated before the first row, as it is when that is requested.
                Assert.IsType(expected, Record.Exception(() => {
                    using var reader = ExcelDocument.OpenReader(path, options); _ = reader.FieldCount;
                }));
            } else {
                using var reader = ExcelDocument.OpenReader(path, options);
                Exception? failure = ReadUntilError(reader, out int rows);
                Assert.IsType(expected, failure);
                if (!malformedXml) Assert.Contains("missing cell style", failure!.Message);
                Assert.True(rows > 0);
            }
        } finally {
            File.Delete(path);
        }
    }

    /// <summary>Reads rows until a read fails and returns that failure with the number of rows returned before it.</summary>
    internal static Exception? ReadUntilError(System.Data.Common.DbDataReader reader, out int rows) {
        int count = 0;
        Exception? failure = Record.Exception(() => {
            while (reader.Read()) count++;
        });
        rows = count;
        return failure;
    }

    [Fact]
    public void OpenDataReader_WindowedWorksheetMatchesTheBufferedReader() {
        // Implicit rows and cells, gaps in row numbers, empty and sparse rows, formulas, entities, line breaks and
        // rows longer than the window, so every row shape crosses window boundaries.
        var rows = new StringBuilder(WindowedIndexHeader);
        int rowNumber = 1;
        for (int index = 0; index < 6000; index++) {
            rowNumber += index % 7 == 0 ? 3 : 1;
            string r = rowNumber.ToString(CultureInfo.InvariantCulture);
            switch (index % 5) {
                case 0:
                    rows.Append("<row r=\"").Append(r).Append("\"><c r=\"A").Append(r).Append("\" t=\"inlineStr\"><is><t>a&lt;b ")
                        .Append(r).Append("</t></is></c><c r=\"C").Append(r).Append("\" s=\"1\"><v>45000.5</v></c></row>");
                    break;
                case 1:
                    rows.Append("<row r=\"").Append(r).Append("\"><c t=\"inlineStr\"><is><t>line\r\nbreak</t></is></c><c><v>")
                        .Append(r).Append("</v></c><c/><c><f>B").Append(r).Append("+1</f><v>").Append(rowNumber + 1).Append("</v></c></row>");
                    break;
                case 2:
                    rows.Append("<row r=\"").Append(r).Append("\"/>");
                    break;
                case 3:
                    string text = index % 1000 == 3 ? new string('x', 100_000) : "plain";
                    rows.Append("<row r=\"").Append(r).Append("\" ht=\"20\" customHeight=\"1\"><c r=\"D").Append(r)
                        .Append("\" t=\"inlineStr\"><is><t>").Append(text).Append("</t></is></c></row>");
                    break;
                default:
                    rows.Append("<row r=\"").Append(r).Append("\"><c r=\"A").Append(r).Append("\" t=\"b\"><v>1</v></c><c r=\"B")
                        .Append(r).Append("\"><v>-0.25</v></c><c r=\"C").Append(r).Append("\" s=\"1\"><v>45001</v></c><c r=\"D")
                        .Append(r).Append("\" t=\"e\"><v>#N/A</v></c></row>");
                    break;
            }
        }
        string path = CreateWindowedIndexWorkbook(WindowedIndexWorksheet(rows.ToString(), dimension: null));
        try {
            List<object?[]> expected = ReadAllValues(path, windowed: false);
            List<object?[]> actual = ReadAllValues(path, windowed: true);
            Assert.True(expected.Count > 6000);
            Assert.Equal(expected.Count, actual.Count);
            for (int index = 0; index < expected.Count; index++) Assert.Equal(expected[index], actual[index]);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("prefixed")]
    [InlineData("leadingEmptyRows")]
    public void OpenDataReader_WindowedWorksheetShapesMatchTheBufferedReader(string shape) {
        var xml = new StringBuilder();
        if (shape == "prefixed") {
            // Prefixed row closing tags are the only row boundaries the window can slide past.
            xml.Append("<x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheetData>");
            for (int row = 1; row <= 4000; row++) {
                string r = row.ToString(CultureInfo.InvariantCulture);
                xml.Append("<x:row r=\"").Append(r).Append("\"><x:c r=\"A").Append(r).Append("\"><x:v>").Append(r)
                    .Append("</x:v></x:c><x:c r=\"B").Append(r).Append("\" t=\"inlineStr\"><x:is><x:t>text ").Append(r)
                    .Append("</x:t></x:is></x:c></x:row>");
            }
            xml.Append("</x:sheetData></x:worksheet>");
        } else {
            // More empty rows than one window holds precede the first cell, so delivery starts past the window.
            xml.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A8001:A9000\"/><sheetData>");
            for (int row = 1; row <= 8000; row++) xml.Append("<row r=\"").Append(row).Append("\" ht=\"20\" customHeight=\"1\"></row>");
            for (int row = 8001; row <= 9000; row++) {
                xml.Append("<row r=\"").Append(row).Append("\"><c r=\"A").Append(row).Append("\"><v>").Append(row).Append("</v></c></row>");
            }
            xml.Append("</sheetData></worksheet>");
        }
        string path = CreateWindowedIndexWorkbook(xml.ToString());
        try {
            List<object?[]> expected = ReadAllValues(path, windowed: false);
            List<object?[]> actual = ReadAllValues(path, windowed: true);
            Assert.Equal(expected.Count, actual.Count);
            for (int index = 0; index < expected.Count; index++) Assert.Equal(expected[index], actual[index]);
            Assert.NotEmpty(expected);
        } finally {
            File.Delete(path);
        }
    }

    private static List<object?[]> ReadAllValues(string path, bool windowed) {
        using var reader = ExcelDocument.OpenReader(path, WorksheetWindowOptions(windowed));
        var rows = new List<object?[]>();
        while (reader.Read()) {
            var values = new object?[reader.FieldCount];
            reader.GetValues(values!);
            rows.Add(values);
        }
        return rows;
    }

    private static string LargeWindowWorksheet(bool dimension, string? lastRowCells, string? tail) {
        int lastRow = LargeWorksheetRows + 1;
        var xml = new StringBuilder(lastRow * 180);
        xml.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        // An extra last-row cell is the fifth column, so the declared range then includes it.
        if (dimension) xml.Append(lastRowCells == null ? "<dimension ref=\"A1:D" : "<dimension ref=\"A1:E").Append(lastRow).Append("\"/>");
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
