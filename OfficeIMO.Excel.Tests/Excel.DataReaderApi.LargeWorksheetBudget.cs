using System.IO.Compression;
using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void DataReader_LargeWorksheetPreservesSparseValuesAndRows(bool declaredDimension, bool prefetch) {
        string path = CreateLargeSparseWorksheet(declaredDimension);
        try {
            using var reader = ExcelDocument.OpenDataReader(path,
                new ExcelReadOptions { EnableWorksheetPrefetch = prefetch });
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("Id", reader.GetName(0));
            Assert.Equal("Value", reader.GetName(1));
            int rows = 0;
            while (reader.Read()) {
                rows++;
                if (rows == 1) {
                    Assert.Equal(42, reader.GetInt32(0));
                    Assert.Equal(7, reader.GetInt32(1));
                } else if (rows == 500001) {
                    Assert.Equal(43, reader.GetInt32(0));
                    Assert.True(reader.IsDBNull(1));
                } else {
                    Assert.True(reader.IsDBNull(0));
                    Assert.True(reader.IsDBNull(1));
                }
            }
            Assert.Equal(500001, rows);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("s=\"99999\"", "<v>43</v>", false, typeof(InvalidDataException))]
    [InlineData("t=\"s\"", "<v>99999</v>", false, typeof(InvalidDataException))]
    [InlineData("", "<f t=\"shared\" si=\"0\"/><v>43</v>", false, typeof(NotSupportedException))]
    [InlineData("", "<v>43</v>", true, typeof(XmlException))]
    public void DataReader_LargeDeclaredWorksheetValidatesBeforeExposingRows(
        string attributes, string content, bool malformedTail, Type error) {
        string path = CreateLargeSparseWorksheet(true, attributes, content, malformedTail);
        try {
            Assert.Throws(error, () => ExcelDocument.OpenDataReader(path,
                new ExcelReadOptions { UseCachedFormulaResult = false }));
        } finally {
            File.Delete(path);
        }
    }

    private static string CreateLargeSparseWorksheet(bool declaredDimension,
        string attributes = "", string content = "<v>43</v>", bool malformedTail = false) {
        string path = CreateCompactFastPathWorkbook();
        try {
            using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
            archive.GetEntry("xl/worksheets/sheet1.xml")!.Delete();
            using var stream = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest).Open();
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            if (declaredDimension) writer.Write("<dimension ref=\"A1:B500002\"/>");
            writer.Write("<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Id</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Value</t></is></c></row>");
            writer.Write("<row r=\"2\"><c r=\"A2\"><v>42</v></c><c r=\"B2\"><v>7</v></c></row>");
            // A small compressed fixture reaches the large-part path without
            // constructing a large string or asserting any host memory/timing budget.
            string padding = new(' ', 64 * 1024);
            for (int block = 0; block < 513; block++) writer.Write(padding);
            writer.Write($"<row r=\"500002\"><c r=\"A500002\" {attributes}>{content}</c></row></sheetData>");
            writer.Write(malformedTail ? "<broken></worksheet>" : "</worksheet>");
            return path;
        } catch {
            File.Delete(path);
            throw;
        }
    }
}
