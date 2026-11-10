using System.Data;
using System.Text;
using System.Xml;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("x", true, false)]
    [InlineData("sheet", true, false)]
    [InlineData("x", false, false)]
    [InlineData("", true, false)]
    [InlineData("x", true, true)]
    public void DataReader_PrefixedWorksheetPreservesTypedValuesAndEmptyCells(string prefix, bool workbookReader, bool dimensionlessImplicitCoordinates) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = PrefixedWorksheetXml(prefix).Replace(" spaced", " literal\r\nline\r spaced");
            if (dimensionlessImplicitCoordinates) {
                var worksheet = System.Xml.Linq.XDocument.Parse(xml);
                var ns = worksheet.Root!.Name.Namespace;
                worksheet.Root.Element(ns + "dimension")!.Remove();
                foreach (var element in worksheet.Descendants(ns + "row").Concat(worksheet.Descendants(ns + "c")))
                    element.Attribute("r")?.Remove();
                var serialized = new StringBuilder();
                using (var writer = XmlWriter.Create(serialized, new XmlWriterSettings {
                    OmitXmlDeclaration = true, NewLineHandling = NewLineHandling.Entitize
                })) {
                    worksheet.Save(writer);
                }
                xml = serialized.ToString();
            }
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var document = workbookReader ? null : ExcelDocumentReader.Open(path);
            using IDataReader reader = workbookReader
                ? ExcelDocument.OpenReader(path)
                : document!.GetSheet("Data").ReadRangeAsDataReader("A1:C3", schemaSampleRows: 0);
            Assert.Equal(3, reader.FieldCount);
            Assert.Equal("Value", reader.GetName(0));
            Assert.Equal("Text", reader.GetName(1));
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
            Assert.Equal(" literal\nline\n spaced\r\n🌍 & <text> ", reader.GetString(1));
            Assert.True(reader.GetBoolean(2));
            Assert.True(reader.Read());
            Assert.Equal(43, reader.GetInt32(0));
            Assert.Equal(string.Empty, reader.GetString(1));
            Assert.True(reader.IsDBNull(2));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("x")]
    public void EnumerateCells_QualifiedWorkbookReturnsFormulaTextCachedValuesDatesAndSharedStrings(string prefix) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.EnumerateQualified.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Name");
                sheet.Cell(1, 2).SetValue(new DateTime(2024, 3, 1));
                document.Save();
            }
            string generated = Encoding.UTF8.GetString(ReadZipEntry(path, "xl/worksheets/sheet1.xml"));
            string dateStyle = System.Text.RegularExpressions.Regex.Match(generated, "<c r=\"B1\" s=\"(\\d+)\"").Groups[1].Value;
            Assert.NotEqual(string.Empty, dateStyle);

            string p = prefix.Length == 0 ? string.Empty : prefix + ":";
            string ns = prefix.Length == 0 ? "xmlns" : "xmlns:" + prefix;
            const string main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            string worksheet = $"<?xml version=\"1.0\" encoding=\"utf-8\"?><{p}worksheet {ns}=\"{main}\"><{p}sheetData>"
                + $"<{p}row r=\"1\"><{p}c r=\"A1\" t=\"s\"><{p}v>0</{p}v></{p}c><{p}c r=\"B1\" s=\"{dateStyle}\"><{p}v>45352</{p}v></{p}c>"
                + $"<{p}c r=\"C1\"><{p}f>A2*2</{p}f></{p}c><{p}c r=\"D1\"><{p}f>1+1</{p}f><{p}v>2</{p}v></{p}c></{p}row>"
                + $"</{p}sheetData></{p}worksheet>";
            string sharedStrings = $"<?xml version=\"1.0\" encoding=\"utf-8\"?><{p}sst {ns}=\"{main}\"><{p}si><{p}t>Name</{p}t></{p}si></{p}sst>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(worksheet));
            ReplaceZipEntry(path, "xl/sharedStrings.xml", Encoding.UTF8.GetBytes(sharedStrings));

            using var loaded = ExcelDocument.Load(new MemoryStream(File.ReadAllBytes(path)));
            ExcelCellValueInfo[] cells = loaded.Sheets[0].EnumerateCells().ToArray();
            Assert.Equal(new[] { 1, 2, 3, 4 }, cells.Select(cell => cell.Column));
            Assert.All(cells, cell => Assert.Equal(1, cell.Row));
            Assert.Equal("Name", cells[0].Value);
            Assert.Equal(new DateTime(2024, 3, 1), cells[1].Value);
            Assert.Equal("A2*2", cells[2].Value);
            Assert.Equal(2d, cells[3].Value);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("self-bound-formula")]
    [InlineData("undeclared-range-widens")]
    public void EnumerateCells_ReturnsEveryCellForWriterShapesOutsideTheCompactLayout(string shape) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.EnumerateShapes.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddSheet("Data").Cell(1, 1).SetValue(1);
                document.Save();
            }
            const string main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            string rows = shape == "self-bound-formula"
                // The Open XML SDK writes a formula created in code this way under a default-namespace root.
                ? $"<row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\" t=\"n\"><x:f ca=\"0\" xmlns:x=\"{main}\">A1*3</x:f><v>6</v></c></row>"
                // No dimension, and the second row starts left of the first.
                : "<row r=\"1\"><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c></row><row r=\"2\"><c r=\"A2\"><v>1</v></c><c r=\"C2\"><v>4</v></c></row>";
            string worksheet = $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"{main}\"><sheetData>{rows}</sheetData></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(worksheet));

            using var loaded = ExcelDocument.Load(new MemoryStream(File.ReadAllBytes(path)));
            string[] Read(bool cachedFormulaResults) => loaded.Sheets[0]
                .EnumerateCells(new ExcelReaderOptions { UseCachedFormulaResult = cachedFormulaResults })
                .Select(cell => $"{ExcelA1.CellReference(cell.Row, cell.Column)}={cell.Value}").ToArray();
            if (shape == "self-bound-formula") {
                Assert.Equal(new[] { "A1=2", "B1=6" }, Read(cachedFormulaResults: true));
                Assert.Equal(new[] { "A1=2", "B1=A1*3" }, Read(cachedFormulaResults: false));
            } else {
                Assert.Equal(new[] { "B1=2", "C1=3", "A2=1", "C2=4" }, Read(cachedFormulaResults: true));
            }
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataReader_PrefixedWorksheetRetainsFallbackForMixedElementPrefixes() {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = PrefixedWorksheetXml("x").Replace("<x:v>43</x:v>",
                "<other:v xmlns:other=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">43</other:v>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var reader = ExcelDocument.OpenReader(path);
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
            Assert.True(reader.Read());
            Assert.Equal(43, reader.GetInt32(0));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("row", "r=\"2\"")]
    [InlineData("c", "r=\"A2\"")]
    [InlineData("v", "")]
    [InlineData("f", "")]
    [InlineData("is", "")]
    [InlineData("t", "xml:space=\"preserve\"")]
    public void DataReader_PrefixedWorksheetRejectsReboundValueNamespaces(string element, string attributes) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = PrefixedWorksheetXml("x");
            string tag = "<x:" + element + (attributes.Length == 0 ? "" : " " + attributes) + ">";
            int start = xml.IndexOf(tag, StringComparison.Ordinal);
            Assert.True(start >= 0);
            xml = xml.Insert(start + tag.Length - 1, " xmlns:x=\"urn:foreign\"");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            Assert.Throws<InvalidDataException>(() => { using var reader = ExcelDocument.OpenReader(path); _ = reader.FieldCount; });
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataReader_PrefixedWorksheetStillValidatesMalformedTrailingXml() {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = PrefixedWorksheetXml("x").Replace("</x:worksheet>", "<!--invalid--comment--></x:worksheet>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            Assert.Throws<XmlException>(() => { using var reader = ExcelDocument.OpenReader(path); _ = reader.FieldCount; });
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenReader_ParallelPrefixedRangePreservesTextAfterSchemaReplay() {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = PrefixedWorksheetXml("x").Replace("<x:is/>",
                "<x:is><x:t xml:space=\"preserve\">A&amp;B / München / 🌍\r\nline\rend&#xD;</x:t></x:is>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions {
                A1Range = "A1:C3", InferSchema = true, SchemaSampleRows = 1
            });
            var rows = reader.RowsAsParallel<PrefixedParallelTextRow>(new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2, BatchSize = 2
            }).ToArray();

            Assert.Equal(new[] { 42, 43 }, rows.Select(row => row.Value));
            Assert.Equal(new[] { " spaced\r\n🌍 & <text> ", "A&B / München / 🌍\nline\nend\r" },
                rows.Select(row => row.Text));
        } finally {
            File.Delete(path);
        }
    }

    public sealed class PrefixedParallelTextRow {
        public int Value { get; set; }
        public string? Text { get; set; }
    }

    private static string PrefixedWorksheetXml(string prefix) {
        string xml = $$"""
        <?xml version="1.0" encoding="utf-8"?>
        <{{prefix}}:worksheet xmlns:{{prefix}}="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <{{prefix}}:dimension ref="A1:C3"/>
          <{{prefix}}:sheetData>
            <{{prefix}}:row r="1"><{{prefix}}:c r="A1" t="inlineStr"><{{prefix}}:is><{{prefix}}:t>Value</{{prefix}}:t></{{prefix}}:is></{{prefix}}:c><{{prefix}}:c r="B1" t="inlineStr"><{{prefix}}:is><{{prefix}}:t>Text</{{prefix}}:t></{{prefix}}:is></{{prefix}}:c><{{prefix}}:c r="C1" t="inlineStr"><{{prefix}}:is><{{prefix}}:t>Enabled</{{prefix}}:t></{{prefix}}:is></{{prefix}}:c></{{prefix}}:row>
            <{{prefix}}:row r="2"><{{prefix}}:c r="A2"><{{prefix}}:f>21*2</{{prefix}}:f><{{prefix}}:v>42</{{prefix}}:v></{{prefix}}:c><{{prefix}}:c r="B2" t="inlineStr"><{{prefix}}:is><{{prefix}}:t xml:space="preserve"> spaced&#xD;&#xA;🌍 &amp; &lt;text&gt; </{{prefix}}:t></{{prefix}}:is></{{prefix}}:c><{{prefix}}:c r="C2" t="b"><{{prefix}}:v>1</{{prefix}}:v></{{prefix}}:c></{{prefix}}:row>
            <{{prefix}}:row r="3"><{{prefix}}:c r="A3"><{{prefix}}:v>43</{{prefix}}:v></{{prefix}}:c><{{prefix}}:c r="B3" t="inlineStr"><{{prefix}}:is/></{{prefix}}:c></{{prefix}}:row>
          </{{prefix}}:sheetData>
        </{{prefix}}:worksheet>
        """;
        return prefix.Length == 0
            ? xml.Replace("<:", "<").Replace("</:", "</").Replace("xmlns:=", "xmlns=")
            : xml;
    }
}
