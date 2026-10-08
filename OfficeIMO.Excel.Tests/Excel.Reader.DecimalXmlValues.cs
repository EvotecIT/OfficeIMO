using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Reader_XmlDecimalValuesPreserveTypesPrecisionAndFormulaPolicy(bool useCachedFormulaResult) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string[] headers = { "Artifact", "Exponent", "Overflow", "Small", "Text", "Formula", "Empty", "Invalid", "Long" };
            string headerXml = string.Concat(headers.Select((header, index) =>
                $"<c r=\"{(char)('A' + index)}1\" t=\"inlineStr\"><is><t>{header}</t></is></c>"));
            string xml = $$"""
                <?xml version="1.0" encoding="utf-16"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                  <row r="1">{{headerXml}}</row>
                  <row r="2">
                    <c r="A2"><v>165258.23999999999</v></c>
                    <c r="B2" t="n"><v>1.2345E+2</v></c>
                    <c r="C2"><v>1E100</v></c>
                    <c r="D2"><v>-0.000001</v></c>
                    <c r="E2" t="str"><v>165258.23999999999</v></c>
                    <c r="F2"><f>1+2</f><v>3</v></c>
                    <c r="G2"><v/></c>
                    <c r="H2"><v>not-a-number</v></c>
                    <c r="I2"><v>{{new string('0', 80)}}12.5</v></c>
                  </row>
                </sheetData></worksheet>
                """;
            // UTF-16 exercises the streaming XML fallback also used beyond the indexed reader's size limit.
            byte[] bytes = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray();
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", bytes);
            var options = new ExcelReadOptions {
                NumericAsDecimal = true,
                UseCachedFormulaResult = useCachedFormulaResult,
                InferDataTableColumnTypes = false
            };
            object[] expected = {
                165258.24m, 123.45m, 1E100, -0.000001m, "165258.23999999999",
                useCachedFormulaResult ? (object)3m : "1+2", string.Empty, "not-a-number", 12.5m
            };

            using (var reader = ExcelDocument.OpenDataReader(path, options)) {
                Assert.Equal(headers.Length, reader.FieldCount);
                Assert.True(reader.Read());
                Assert.Equal(165258.24m, reader.GetDecimal(0));
                Assert.Equal(165258.24m, Assert.IsType<decimal>(reader.GetValue(0)));
                Assert.Equal(165258.24d, reader.GetDouble(0));
                Assert.Equal(1E100, reader.GetDouble(2));
                for (int column = 0; column < expected.Length; column++) {
                    Assert.Equal(headers[column], reader.GetName(column));
                    Assert.Equal(expected[column].GetType(), reader.GetValue(column).GetType());
                    Assert.Equal(expected[column], reader.GetValue(column));
                }
                Assert.False(reader.Read());
            }
            using (var owner = ExcelDocumentReader.Open(path, options)) {
                object?[,] values = owner.GetSheet("Data").ReadRange("A2:I2", ExcelExecutionMode.Sequential);
                for (int column = 0; column < expected.Length; column++) Assert.Equal(expected[column], values[0, column]);
            }
            using (var owner = ExcelDocumentReader.Open(path, options))
            using (var table = owner.GetSheet("Data").ReadRangeAsDataTable("A1:I2", headersInFirstRow: true)) {
                Assert.Single(table.Rows.Cast<System.Data.DataRow>());
                for (int column = 0; column < expected.Length; column++) Assert.Equal(expected[column], table.Rows[0][column]);
            }
        } finally {
            File.Delete(path);
        }
    }
}
