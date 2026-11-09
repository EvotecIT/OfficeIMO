using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DataReader_ImplicitRowFallbackKeepsExplicitRowAndUsedBounds(bool startsInSecondColumn) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string firstColumn = startsInSecondColumn ? "B" : "A";
            string secondColumn = startsInSecondColumn ? "C" : "B";
            string provisionalFirstCell = startsInSecondColumn
                ? "<c r=\"B5\"><v>50</v></c>"
                : "<c><v>50</v></c>";
            string xml = $$"""
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                  <row r="3"><c r="{{firstColumn}}3"><v>30</v></c><c r="{{secondColumn}}3"><v>31</v></c></row>
                  <row>{{provisionalFirstCell}}<c r="{{secondColumn}}5"><v>51</v></c></row>
                  <row r="6"><c r="{{firstColumn}}6"><v>60</v></c><c r="{{secondColumn}}6"><v>61</v></c></row>
                </sheetData></worksheet>
                """;
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            using var reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions { HasHeaderRow = false });
            Assert.Equal(2, reader.FieldCount);
            Assert.True(reader.Read());
            Assert.Equal(30D, reader.GetValue(0));
            Assert.Equal(31D, reader.GetValue(1));
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.True(reader.IsDBNull(1));
            Assert.True(reader.Read());
            Assert.Equal(50D, reader.GetValue(0));
            Assert.Equal(51D, reader.GetValue(1));
            Assert.True(reader.Read());
            Assert.Equal(60D, reader.GetValue(0));
            Assert.Equal(61D, reader.GetValue(1));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void DataReader_ImplicitRowValidationReportsTheResolvedCell(bool prefixed, bool sharedString, bool cachedRow) {
        string path = CreateCompactFastPathWorkbook();
        try {
            if (sharedString) {
                using var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, true);
                var workbook = package.WorkbookPart!;
                var strings = workbook.SharedStringTablePart
                    ?? workbook.AddNewPart<DocumentFormat.OpenXml.Packaging.SharedStringTablePart>();
                strings.SharedStringTable = new DocumentFormat.OpenXml.Spreadsheet.SharedStringTable(
                    new DocumentFormat.OpenXml.Spreadsheet.SharedStringItem(
                        new DocumentFormat.OpenXml.Spreadsheet.Text("valid")));
                strings.SharedStringTable.Save();
            }
            string attributes = sharedString ? "t=\"s\"" : "s=\"99999\"";
            string value = sharedString ? "99999" : "42";
            string prime = cachedRow ? "<row><c t=\"s\"><v>0</v></c><c><v>1</v></c></row>" : string.Empty;
            string xml = $$"""
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Value</t></is></c><c r="B1" t="inlineStr"><is><t>Other</t></is></c></row>{{prime}}<row><c {{attributes}}><v>{{value}}</v></c><c r="B10"><v>43</v></c></row></sheetData></worksheet>
                """;
            if (prefixed) xml = ApplyDimensionlessWorksheetPrefix(xml);
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            var error = Assert.Throws<InvalidDataException>(() => {
                using var reader = ExcelDocument.OpenReader(path);
                _ = reader.FieldCount;
            });
            string missing = sharedString ? "shared string" : "cell style";
            Assert.Equal($"Worksheet 'Data' cell A10 references a missing {missing}.", error.Message);
        } finally {
            File.Delete(path);
        }
    }

}
