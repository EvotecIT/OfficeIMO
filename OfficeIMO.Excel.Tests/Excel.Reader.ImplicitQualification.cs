using System.Data;
using System.Text;
using System.Threading;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void DataReader_ImplicitDimensionlessRowsKeepEmptyRowPositionsAndIgnoreCommentTokens() {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = ImplicitQualificationWorksheet("<row/>" + ImplicitQualificationHeader
                + "<row><c><v>42</v></c><c t=\"inlineStr\"><is><t>Alpha</t></is></c><c><v>7</v></c></row>"
                + "<!-- producer note: <row><c><v>999</v></c></row> -->"
                + "<row/>"
                + "<row><c><v>43</v></c><c t=\"inlineStr\"><is><t>Beta</t></is></c><c><v>8</v></c></row>"
                + "<row/>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            using var reader = ExcelDocument.OpenReader(path);
            Assert.Equal(3, reader.FieldCount);
            Assert.Equal("Id", reader.GetName(0));
            AssertImplicitQualificationRow(reader, 42, "Alpha", 7);
            Assert.True(reader.Read());
            for (int column = 0; column < reader.FieldCount; column++) Assert.True(reader.IsDBNull(column));
            AssertImplicitQualificationRow(reader, 43, "Beta", 8);
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataReader_ImplicitDimensionlessRowsPreserveChangingWidthsAndEmptyRows() {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = ImplicitQualificationWorksheet(
                "<row/>"
                + ImplicitQualificationHeader
                + ImplicitQualificationRows
                + "<row/>"
                + "<row customFormat=\"1\"><c><v>44</v></c><c t=\"inlineStr\"><is><t>Gamma</t></is></c><c><v>9</v></c><c><v>99</v></c></row>"
                + "<row><c><v>45</v></c></row>"
                + "<row/>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            using var reader = ExcelDocument.OpenReader(path);
            Assert.Equal(4, reader.FieldCount);
            Assert.Equal("Id", reader.GetName(0));
            Assert.Equal("Name", reader.GetName(1));
            Assert.Equal("Value", reader.GetName(2));
            AssertImplicitQualificationRow(reader, 42, "Alpha", 7);
            Assert.True(reader.IsDBNull(3));
            AssertImplicitQualificationRow(reader, 43, "Beta", 8);
            Assert.True(reader.IsDBNull(3));
            Assert.True(reader.Read());
            for (int column = 0; column < reader.FieldCount; column++) Assert.True(reader.IsDBNull(column));
            AssertImplicitQualificationRow(reader, 44, "Gamma", 9);
            Assert.Equal(99, reader.GetInt32(3));
            Assert.True(reader.Read());
            Assert.Equal(45, reader.GetInt32(0));
            for (int column = 1; column < reader.FieldCount; column++) Assert.True(reader.IsDBNull(column));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DataReader_ImplicitRowsClearReusedValueAndFormulaMetadata(bool useCachedFormulaResult) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string formulas = "<row><c><f>1+1</f><v>2</v></c><c><f>2+2</f><v>4</v></c><c><f>3+3</f><v>6</v></c></row>";
            string xml = ImplicitQualificationWorksheet(formulas
                + "<row><c><v>10</v></c><c/><c></c></row>"
                + "<row><c><v>30</v></c><c><f>20+2</f><v>22</v></c><c><v>31</v></c></row>"
                + "<row><c><v>11</v></c></row><row></row>"
                + "<row><c><v>12</v></c><c t=\"inlineStr\"><is><t></t></is></c><c><v>13</v></c></row>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            using var reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions {
                HasHeaderRow = false, MaxDataReaderBufferedCells = 3, MaxDataReaderChunkRows = 1,
                UseCachedFormulaResult = useCachedFormulaResult,
            });
            Assert.Equal(3, reader.FieldCount);
            Assert.True(reader.Read());
            for (int column = 0; column < 3; column++) {
                if (useCachedFormulaResult) Assert.Equal((column + 1) * 2, reader.GetInt32(column));
                else Assert.Equal($"{column + 1}+{column + 1}", reader.GetString(column));
            }
            Assert.True(reader.Read());
            Assert.Equal(10, reader.GetInt32(0));
            Assert.True(reader.IsDBNull(1));
            Assert.True(reader.IsDBNull(2));
            Assert.True(reader.Read());
            Assert.Equal(30, reader.GetInt32(0));
            if (useCachedFormulaResult) Assert.Equal(22, reader.GetInt32(1));
            else Assert.Equal("20+2", reader.GetString(1));
            Assert.Equal(31, reader.GetInt32(2));
            Assert.True(reader.Read());
            Assert.Equal(11, reader.GetInt32(0));
            Assert.True(reader.IsDBNull(1));
            Assert.True(reader.IsDBNull(2));
            Assert.True(reader.Read());
            for (int column = 0; column < reader.FieldCount; column++) Assert.True(reader.IsDBNull(column));
            Assert.True(reader.Read());
            Assert.Equal(12, reader.GetInt32(0));
            Assert.Equal(string.Empty, reader.GetString(1));
            Assert.Equal(13, reader.GetInt32(2));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DataReader_ImplicitDimensionlessRowsInferLateCellReferencesBeforeEarlierCells(bool includeDimension) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = ImplicitQualificationWorksheet(
                ImplicitQualificationHeader + ImplicitQualificationRows
                + "<row><c><v>44</v></c><c r=\"B6\" t=\"inlineStr\"><is><t>Gamma</t></is></c><c><v>9</v></c></row>"
                + "<row><c><v>45</v></c><c t=\"inlineStr\"><is><t>Delta</t></is></c><c><v>10</v></c></row>");
            if (includeDimension) xml = xml.Replace("<sheetData>", "<dimension ref=\"A1:C7\"/><sheetData>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            using var reader = ExcelDocument.OpenReader(path);
            Assert.Equal(3, reader.FieldCount);
            AssertImplicitQualificationRow(reader, 42, "Alpha", 7);
            AssertImplicitQualificationRow(reader, 43, "Beta", 8);
            for (int row = 4; row <= 5; row++) {
                Assert.True(reader.Read());
                for (int column = 0; column < reader.FieldCount; column++) Assert.True(reader.IsDBNull(column));
            }
            AssertImplicitQualificationRow(reader, 44, "Gamma", 9);
            AssertImplicitQualificationRow(reader, 45, "Delta", 10);
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("footer")]
    [InlineData("comment")]
    [InlineData("duplicate-attribute")]
    [InlineData("unknown-entity")]
    [InlineData("forbidden-scalar")]
    public void DataReader_ImplicitDimensionlessRowsRejectMalformedXmlBeforeDelivery(string failure) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string lateRow = "<row><c><v>44</v></c><c t=\"inlineStr\"><is><t>Gamma</t></is></c><c><v>9</v></c></row>";
            string footer = string.Empty;
            switch (failure) {
                case "footer": footer = "<broken>"; break;
                case "comment": lateRow = "<!--invalid--comment-->" + lateRow; break;
                case "duplicate-attribute": lateRow = lateRow.Replace("t=\"inlineStr\"", "t=\"inlineStr\" t=\"inlineStr\""); break;
                case "unknown-entity": lateRow = lateRow.Replace("Gamma", "&unknown;"); break;
                case "forbidden-scalar": lateRow = lateRow.Replace("Gamma", "\uffff"); break;
            }
            string xml = ImplicitQualificationWorksheet(ImplicitQualificationHeader + ImplicitQualificationRows + lateRow, footer);
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            Assert.Throws<XmlException>(() => { using var reader = ExcelDocument.OpenReader(path); _ = reader.FieldCount; });
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("<v xmlns=\"urn:foreign\">44</v>")]
    [InlineData("<is><t xmlns=\"urn:foreign\">Gamma</t></is>")]
    public void DataReader_ImplicitDimensionlessRowsRejectForeignValueNamespacesBeforeDelivery(string value) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string cell = value.StartsWith("<is>", StringComparison.Ordinal)
                ? "<c t=\"inlineStr\">" + value + "</c>"
                : "<c>" + value + "</c>";
            string xml = ImplicitQualificationWorksheet(ImplicitQualificationHeader + ImplicitQualificationRows + "<row>" + cell + "</row>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            Assert.Throws<InvalidDataException>(() => { using var reader = ExcelDocument.OpenReader(path); _ = reader.FieldCount; });
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("s=\"99999\"", "<v>44</v>", typeof(InvalidDataException))]
    [InlineData("t=\"s\"", "<v>99999</v>", typeof(InvalidDataException))]
    [InlineData("", "<f t=\"shared\" si=\"0\"/><v>44</v>", typeof(NotSupportedException))]
    public void DataReader_ImplicitDimensionlessRowsValidateLateMetadataBeforeDelivery(string attributes, string content, Type error) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = ImplicitQualificationWorksheet(ImplicitQualificationHeader + ImplicitQualificationRows
                + "<row><c " + attributes + ">" + content + "</c></row>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            Assert.Throws(error, () => { using var reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions { UseCachedFormulaResult = false }); _ = reader.FieldCount; });
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DataReader_ImplicitDimensionlessRowsHonorDataReaderBudgets(bool columnBudget) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = ImplicitQualificationWorksheet(ImplicitQualificationHeader + ImplicitQualificationRows);
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            var options = new ExcelReaderOptions();
            if (columnBudget) options.MaxDataReaderColumns = 2;
            else options.MaxDataReaderBufferedCells = 2;

            var error = Assert.Throws<OfficeInspectionResourceLimitException>(() => { using var reader = ExcelDocument.OpenReader(path, options); _ = reader.FieldCount; });
            Assert.Equal("excel", error.Owner); Assert.Equal(2, error.Maximum);
            Assert.Equal(columnBudget ? nameof(ExcelReaderOptions.MaxDataReaderColumns) : nameof(ExcelReaderOptions.MaxDataReaderBufferedCells), error.LimitName);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataReader_ImplicitDimensionlessRowsRejectColumnsOutsideExcelGrid() {
        string path = CreateCompactFastPathWorkbook();
        try {
            var row = new StringBuilder("<row>");
            for (int column = 0; column <= 16_384; column++) row.Append("<c><v>1</v></c>");
            row.Append("</row>");
            string xml = ImplicitQualificationWorksheet(row.ToString());
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            Assert.Throws<ArgumentException>(() => { using var reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions { HasHeaderRow = false }); _ = reader.FieldCount; });
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataReader_ImplicitDimensionlessRowsRejectOverstatedPackageLengthBeforeDelivery() {
        string path = CreateCompactFastPathWorkbook();
        try {
            byte[] xml = Encoding.UTF8.GetBytes(ImplicitQualificationWorksheet(ImplicitQualificationHeader + ImplicitQualificationRows));
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", xml);
            byte[] package = File.ReadAllBytes(path);
            OpenXmlPartLengthTests.SetDeclaredLength(package, "xl/worksheets/sheet1.xml", xml.Length + 1);

            Assert.Throws<InvalidDataException>(() => { using var reader = ExcelDocument.OpenReader(package); _ = reader.FieldCount; });
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataReader_ImplicitDimensionlessRowsObserveCancellationDuringTraversal() {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = ImplicitQualificationWorksheet(ImplicitQualificationHeader + ImplicitQualificationRows);
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var cancellation = new CancellationTokenSource();
            using var reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions { CancellationToken = cancellation.Token });
            Assert.True(reader.Read());
            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() => reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    private const string ImplicitQualificationHeader = "<row><c t=\"inlineStr\"><is><t>Id</t></is></c><c t=\"inlineStr\"><is><t>Name</t></is></c><c t=\"inlineStr\"><is><t>Value</t></is></c></row>";
    private const string ImplicitQualificationRows = "<row><c><v>42</v></c><c t=\"inlineStr\"><is><t>Alpha</t></is></c><c><v>7</v></c></row>"
        + "<row><c><v>43</v></c><c t=\"inlineStr\"><is><t>Beta</t></is></c><c><v>8</v></c></row>";

    private static string ImplicitQualificationWorksheet(string rows, string footer = "") =>
        "<?xml version=\"1.0\" encoding=\"utf-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
        + rows + "</sheetData>" + footer + "</worksheet>";

    private static void AssertImplicitQualificationRow(IDataReader reader, int id, string name, int value) {
        Assert.True(reader.Read());
        Assert.Equal(id, reader.GetInt32(0));
        Assert.Equal(name, reader.GetString(1));
        Assert.Equal(value, reader.GetInt32(2));
    }
}
