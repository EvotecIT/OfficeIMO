using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Data;
using OfficeIMO.Excel;
using System.Data.Common;
using System.Globalization;
using System.Xml;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred, "[h]:mm")]
    [InlineData(ExcelDateSystem.NineteenFour, "[h]:mm")]
    [InlineData(ExcelDateSystem.NineteenHundred, "yyyy-mm-dd")]
    [InlineData(ExcelDateSystem.NineteenFour, "yyyy-mm-dd")]
    public void Reader_NumericReplay_PreservesSampledAndUnsampledSerials(ExcelDateSystem system, string format) {
        string path = CreateNumericReplayWorkbook(system, format);
        for (int api = 0; api < 4; api++) {
            var options = new ExcelReadOptions { SheetName = "Data", InferSchema = true, SchemaSampleRows = 1,
                TypeConverter = (value, _, _) => { Assert.IsType<DateTime>(value); return (false, null); } };
            using var reader = ExcelDocument.OpenDataReader(path, options);
            Assert.Equal(typeof(DateTime), reader.GetFieldType(0));
            Action<RowMapper<NumericFallbackRow>> map = mapper => mapper.FromColumn<double?>("Value", (row, value) => { row.Value = value; return row; });
            var parallel = new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 };
            var rows = api switch {
                0 => reader.RowsAs<NumericFallbackRow>().ToArray(),
                1 => reader.RowsAsParallel<NumericFallbackRow>(parallel).ToArray(),
                2 => reader.RowsAs(map).ToArray(),
                _ => reader.RowsAsParallel(map, parallel).ToArray()
            };
            Assert.Equal(new double?[] { 1.5d, 60.5d }, rows.Select(row => row.Value).ToArray());
        }
    }

    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred)]
    [InlineData(ExcelDateSystem.NineteenFour)]
    public void Reader_NumericReplay_PreservesUnsortedXmlSerials(ExcelDateSystem system) {
        string path = CreateNumericReplayWorkbook(system, "[h]:mm");
        using (var package = SpreadsheetDocument.Open(path, true)) {
            var part = package.WorkbookPart!.WorksheetParts.Single();
            var data = part.Worksheet.GetFirstChild<SheetData>()!;
            var rows = data.Elements<Row>().ToArray();
            data.RemoveAllChildren<Row>();
            data.Append(rows[0], rows[2], rows[1]);
            part.Worksheet.Save();
        }
        int calls = 0;
        using var document = ExcelDocumentReader.Open(path, new ExcelReadOptions {
            CellValueConverter = context => { if (context.RawText is "1.5" or "60.5") calls++; return ExcelCellValue.NotHandled; }
        });
        using var reader = (DbDataReader)document.GetSheet("Data").ReadRangeAsDataReader("A1:A4098", schemaSampleRows: 0);
        var mapped = reader.RowsAs<NumericFallbackRow>().Take(2).ToArray();
        Assert.Equal(new double?[] { 1.5d, 60.5d }, mapped.Select(row => row.Value).ToArray());
        Assert.Equal(2, calls);
    }

    [Theory]
    [InlineData(1, 1)]
    [InlineData(0, 1)]
    [InlineData(1, 3)]
    [InlineData(0, 3)]
    [InlineData(1, 8)]
    [InlineData(0, 8)]
    [InlineData(1, 10)]
    [InlineData(0, 10)]
    public void Reader_NumericReplay_MaterializesDateHeaders(int samples, int width) {
        string path = CreateNumericReplayWorkbook(ExcelDateSystem.NineteenFour, "[h]:mm");
        using (var package = SpreadsheetDocument.Open(path, true)) {
            var part = package.WorkbookPart!.WorksheetParts.Single();
            var cell = part.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First().Elements<Cell>().Single();
            cell.DataType = null;
            cell.CellValue = new CellValue("1.5");
            cell.StyleIndex = part.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Skip(1).First().Elements<Cell>().Single().StyleIndex?.Value;
            part.Worksheet.Save();
        }
        using var document = ExcelDocumentReader.Open(path, new ExcelReadOptions { NormalizeHeaders = false, CellValueConverter = _ => ExcelCellValue.NotHandled });
        using var reader = document.GetSheet("Data").ReadRangeAsDataReader($"A1:{(char)('A' + width - 1)}4098", schemaSampleRows: samples);
        Assert.Equal(DateTime.FromOADate(1.5d).ToString(), reader.GetName(0));
        Assert.True(reader.Read());
        Assert.Equal(1.5d, reader.GetDouble(0));
    }

    private sealed class DuplicateSerialRow {
        public DateTime Date { get; set; }
        public double Serial { get; set; }
    }

    [Fact]
    public void Reader_NumericReplay_MapsTheSameColumnToDateAndSerial() {
        string path = CreateNumericReplayWorkbook(ExcelDateSystem.NineteenFour, "[h]:mm");
        foreach (bool numericFirst in new[] { false, true }) {
            foreach (bool parallel in new[] { false, true }) {
                using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                    InferSchema = true, SchemaSampleRows = 1,
                    TypeConverter = (value, _, _) => { Assert.IsType<DateTime>(value); return (false, null); }
                });
                Action<RowMapper<DuplicateSerialRow>> configure = mapper => {
                    void Numeric() => mapper.FromColumn<double>("Value", (row, value) => { row.Serial = value; return row; });
                    void Date() => mapper.FromColumn<DateTime>("Value", (row, value) => { row.Date = value; return row; });
                    if (numericFirst) { Numeric(); Date(); } else { Date(); Numeric(); }
                };
                var rows = (parallel
                    ? reader.RowsAsParallel(configure, new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 })
                    : reader.RowsAs(configure)).ToArray();
                Assert.Equal(new[] { 1.5d, 60.5d }, rows.Select(row => row.Serial).ToArray());
                Assert.Equal(new[] { DateTime.FromOADate(1.5d), DateTime.FromOADate(60.5d) }, rows.Select(row => row.Date).ToArray());
            }
        }
    }

    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred, "[h]:mm")]
    [InlineData(ExcelDateSystem.NineteenFour, "[h]:mm")]
    [InlineData(ExcelDateSystem.NineteenHundred, "yyyy-mm-dd")]
    [InlineData(ExcelDateSystem.NineteenFour, "yyyy-mm-dd")]
    public void Reader_NumericReplay_PreservesXmlDateFirstGetterOrder(ExcelDateSystem system, string format) {
        string path = CreateNumericReplayWorkbook(system, format);
        using (var package = SpreadsheetDocument.Open(path, true)) {
            var part = package.WorkbookPart!.WorksheetParts.Single();
            var xml = new XmlDocument();
            using (var stream = part.GetStream(FileMode.Open, FileAccess.Read)) xml.Load(stream);
            foreach (XmlElement element in xml.SelectNodes("//*")!) {
                if (element.NamespaceURI == "http://schemas.openxmlformats.org/spreadsheetml/2006/main") element.Prefix = "q";
            }
            xml.DocumentElement!.SetAttribute("xmlns:q", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            using var target = part.GetStream(FileMode.Create, FileAccess.Write);
            xml.Save(target);
        }
        using var document = ExcelDocumentReader.Open(path);
        using var reader = document.GetSheet("Data").ReadRangeAsDataReader("A1:A4098", schemaSampleRows: 0);
        foreach (double serial in new[] { 1.5d, 60.5d }) {
            Assert.True(reader.Read());
            DateTime expected = format == "[h]:mm" ? DateTime.FromOADate(serial) : ExcelDateSystemConverter.FromSerial(serial, system);
            Assert.Equal(expected, reader.GetDateTime(0));
            Assert.Equal(expected, reader.GetValue(0));
            Assert.Equal(serial, reader.GetDouble(0));
        }
    }

    private string CreateNumericReplayWorkbook(ExcelDateSystem system, string format) {
        string path = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
        using var document = ExcelDocument.Create(path);
        document.DateSystem = system;
        var sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Value");
        sheet.CellValue(2, 1, 1.5d);
        sheet.CellValue(3, 1, 60.5d);
        sheet.ColumnStyleByHeader("Value").NumberFormat(format);
        document.Save();
        return path;
    }
}
