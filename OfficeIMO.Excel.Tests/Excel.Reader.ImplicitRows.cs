using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    public static IEnumerable<object[]> ImplicitRowReadCases() {
        foreach (string api in new[] { "OpenDataReader", "BufferedDataReader", "StreamingDataReader", "Range", "Rows", "DataTable", "Objects", "Cells", "Stream" }) {
            foreach (bool utf16 in new[] { false, true })
                foreach (bool prefixed in new[] { false, true })
                    yield return new object[] { api, utf16, prefixed, "Xml" };
            yield return new object[] { api, false, false, "Dom" };
            yield return new object[] { api, false, false, "MissingCellsDom" };
            yield return new object[] { api, false, true, "MissingCells" };
            yield return new object[] { api, true, true, "MissingCells" };
        }
        foreach (string api in new[] { "StreamingDataReader", "BufferedDataReader", "Range" })
            foreach (bool utf16 in new[] { false, true })
                yield return new object[] { api, utf16, true, "Unsorted" };
        foreach (string api in new[] { "OpenDataReader", "Range" })
            foreach (bool utf16 in new[] { false, true })
                yield return new object[] { api, utf16, true, "EscapedReference" };
        foreach (string api in new[] { "Column", "EnumerateRange", "Dictionaries" }) {
            yield return new object[] { api, false, false, "MissingCellsDom" };
            yield return new object[] { api, false, true, "MissingCells" };
            yield return new object[] { api, true, true, "MissingCells" };
        }
    }

    [Theory]
    [MemberData(nameof(ImplicitRowReadCases))]
    public void Reader_ImplicitRowsRetainCellReferencedPositions(string api, bool utf16, bool prefixed, string layout) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string p = prefixed ? "s:" : string.Empty;
            string ns = prefixed ? "xmlns:s" : "xmlns";
            string xml = $$"""
                <{{p}}worksheet {{ns}}="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><{{p}}sheetData>
                  <{{p}}row r="2"/>
                  <{{p}}row r="3"><{{p}}c r="B3" t="inlineStr"><{{p}}is><{{p}}t>Id</{{p}}t></{{p}}is></{{p}}c><{{p}}c r="C3" t="inlineStr"><{{p}}is><{{p}}t>Name</{{p}}t></{{p}}is></{{p}}c><{{p}}c r="D3" t="inlineStr"><{{p}}is><{{p}}t>Amount</{{p}}t></{{p}}is></{{p}}c></{{p}}row>
                  <{{p}}row><{{p}}c r="B4"><{{p}}v>42</{{p}}v></{{p}}c><{{p}}c r="C4" t="str"><{{p}}v>Alpha</{{p}}v></{{p}}c><{{p}}c r="D4"><{{p}}v>7</{{p}}v></{{p}}c></{{p}}row>
                  <{{p}}row><{{p}}c r="B6"><{{p}}v>43</{{p}}v></{{p}}c><{{p}}c r="C6" t="str"><{{p}}v>Beta</{{p}}v></{{p}}c><{{p}}c r="D6"><{{p}}v>9</{{p}}v></{{p}}c></{{p}}row>
                  <{{p}}row><{{p}}c r="B7"><{{p}}v>44</{{p}}v></{{p}}c><{{p}}c r="C7" t="str"><{{p}}v>Gamma</{{p}}v></{{p}}c><{{p}}c r="D7"><{{p}}v>11</{{p}}v></{{p}}c></{{p}}row>
                  <{{p}}row r="10"/>
                </{{p}}sheetData></{{p}}worksheet>
                """;
            if (layout == "Unsorted") {
                var document = System.Xml.Linq.XDocument.Parse(xml);
                var rows = document.Root!.Elements().Single().Elements().ToArray();
                rows[3].Remove();
                rows[2].AddBeforeSelf(rows[3]);
                xml = document.ToString();
            } else if (layout == "EscapedReference") {
                xml = xml.Replace("r=\"B6\"", "r=\"&#66;6\"");
            }
            bool missingCells = layout.StartsWith("MissingCells", StringComparison.Ordinal);
            if (missingCells) {
                var document = System.Xml.Linq.XDocument.Parse(xml);
                var rows = document.Root!.Elements().Single().Elements().ToArray();
                foreach (var cell in rows.SelectMany(row => row.Elements())) {
                    string reference = cell.Attribute("r")!.Value;
                    cell.SetAttributeValue("r", (char)(reference[0] - 1) + reference.Substring(1));
                }
                // Infer A before the B reference, C after it, and the entire final row.
                foreach (var row in rows.Skip(2).Take(3)) {
                    var cells = row.Elements().ToArray();
                    cells[0].Attribute("r")!.Remove();
                    cells[2].Attribute("r")!.Remove();
                }
                rows[3].SetAttributeValue("r", "6");
                rows[4].Elements().ElementAt(1).Attribute("r")!.Remove();
                xml = document.ToString();
            }
            byte[] bytes = utf16 ? Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray() : Encoding.UTF8.GetBytes(xml);
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", bytes);
            var options = new ExcelReadOptions { NumericAsDecimal = true, InferDataTableColumnTypes = false };
            if (layout.EndsWith("Dom", StringComparison.Ordinal)) options.Culture = System.Globalization.CultureInfo.GetCultureInfo("fr-FR");
            object?[,] expected = { { 42m, "Alpha", 7m }, { null, null, null }, { 43m, "Beta", 9m }, { 44m, "Gamma", 11m } };
            using var owner = ExcelDocumentReader.Open(path, options);
            var sheet = owner.GetSheet("Data");
            string dataRange = missingCells ? "A4:C7" : "B4:D7";
            string usedRange = missingCells ? "A3:C7" : "B3:D7";
            Assert.Equal(usedRange, sheet.GetUsedRangeA1());
            object?[,] actual = new object?[4, 3];
            switch (api) {
                case "OpenDataReader":
                case "BufferedDataReader":
                case "StreamingDataReader": {
                    using var reader = api == "OpenDataReader" ? ExcelDocument.OpenDataReader(path, options)
                        : sheet.ReadRangeAsDataReader(api == "StreamingDataReader" ? (missingCells ? "A4:C4100" : "B4:D4100") : dataRange, headersInFirstRow: false, schemaSampleRows: 0);
                    for (int row = 0; row < 4; row++) {
                        Assert.True(reader.Read());
                        for (int col = 0; col < 3; col++) actual[row, col] = reader.IsDBNull(col) ? null : reader.GetValue(col);
                    }
                    if (api == "StreamingDataReader") {
                        for (int row = 8; row <= 4100; row++) {
                            Assert.True(reader.Read());
                            for (int col = 0; col < 3; col++) Assert.True(reader.IsDBNull(col));
                        }
                    }
                    Assert.False(reader.Read());
                    break;
                }
                case "Range": actual = sheet.ReadRange(dataRange); break;
                case "Rows": {
                    var rows = sheet.ReadRows(dataRange).ToArray();
                    Assert.Equal(4, rows.Length);
                    for (int row = 0; row < rows.Length; row++)
                        for (int col = 0; col < 3; col++) actual[row, col] = rows[row]?[col];
                    break;
                }
                case "DataTable": {
                    using var table = sheet.ReadRangeAsDataTable(usedRange, headersInFirstRow: true);
                    Assert.Equal(4, table.Rows.Count);
                    for (int row = 0; row < 4; row++)
                        for (int col = 0; col < 3; col++) actual[row, col] = table.Rows[row].IsNull(col) ? null : table.Rows[row][col];
                    break;
                }
                case "Objects": {
                    var rows = sheet.ReadObjects<ImplicitRowRecord>(usedRange).ToArray();
                    Assert.Equal(4, rows.Length);
                    for (int row = 0; row < 4; row++) {
                        actual[row, 0] = rows[row].Id;
                        actual[row, 1] = rows[row].Name;
                        actual[row, 2] = rows[row].Amount;
                    }
                    break;
                }
                case "Cells":
                    foreach (var cell in sheet.EnumerateCells().Where(cell => cell.Row >= 4)) actual[cell.Row - 4, cell.Column - (missingCells ? 1 : 2)] = cell.Value;
                    break;
                case "EnumerateRange":
                    foreach (var cell in sheet.EnumerateRange(dataRange)) actual[cell.Row - 4, cell.Column - 1] = cell.Value;
                    break;
                case "Column":
                    for (int col = 0; col < 3; col++) {
                        char name = (char)('A' + col);
                        var column = sheet.ReadColumn($"{name}4:{name}7").ToArray();
                        Assert.Equal(4, column.Length);
                        for (int row = 0; row < 4; row++) actual[row, col] = column[row];
                    }
                    break;
                case "Dictionaries": {
                    var rows = sheet.ReadObjects(usedRange).ToArray();
                    Assert.Equal(4, rows.Length);
                    string[] names = { "Id", "Name", "Amount" };
                    for (int row = 0; row < 4; row++)
                        for (int col = 0; col < 3; col++) actual[row, col] = rows[row][names[col]];
                    break;
                }
                case "Stream":
                    foreach (var chunk in sheet.ReadRangeStream(dataRange, chunkRows: 2))
                        for (int row = 0; row < chunk.RowCount; row++)
                            for (int col = 0; col < 3; col++) actual[chunk.StartRow + row - 4, col] = chunk.Rows[row][col];
                    break;
            }
            for (int row = 0; row < 4; row++)
                for (int col = 0; col < 3; col++) Assert.Equal(expected[row, col], actual[row, col]);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void Reader_ImplicitLiveRowProjectionPreservesWorkbookModel() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Data");
        sheet.CellValue(10, 1, "First");
        sheet.CellValue(10, 2, "Value");
        var worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet;
        var row = worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().Single();
        row.RowIndex = null;
        row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().First().CellReference = null;
        string before = worksheet.OuterXml;

        var cells = sheet.EnumerateCells().ToArray();

        Assert.Equal(2, cells.Length);
        Assert.Equal(10, cells[0].Row);
        Assert.Equal(1, cells[0].Column);
        Assert.Equal("First", cells[0].Value);
        var cell = cells[1];
        Assert.Equal(10, cell.Row);
        Assert.Equal(2, cell.Column);
        Assert.Equal("Value", cell.Value);
        Assert.Null(row.RowIndex);
        Assert.Equal(before, worksheet.OuterXml);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Reader_ImplicitLiveRowsRetainHeaderMapAndWorkbookModel(bool omitFirstCellReference, bool unsorted) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Data");
        sheet.CellValue(6, 1, "Id");
        sheet.CellValue(6, 2, "Name");
        sheet.CellValue(7, 1, 42);
        sheet.CellValue(7, 2, "Alpha");
        var worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet;
        foreach (var row in worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>()) {
            row.RowIndex = null;
            if (omitFirstCellReference) row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().First().CellReference = null;
        }
        if (unsorted) {
            var data = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()!;
            var row = data.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Last();
            row.Remove();
            data.PrependChild(row);
        }
        string before = worksheet.OuterXml;

        var headers = sheet.GetHeaderMap();

        Assert.Equal(2, headers.Count);
        Assert.Equal(1, headers["Id"]);
        Assert.Equal(2, headers["Name"]);
        Assert.Equal(headers, sheet.GetHeaderMap());
        Assert.Equal(before, worksheet.OuterXml);

        var headerRow = worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().Single(row =>
            row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Any(cell => cell.CellReference?.Value == "B6"));
        headerRow.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell {
            CellReference = "C6",
            DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Extra")
        });
        string updated = worksheet.OuterXml;
        var refreshed = sheet.GetHeaderMap();
        Assert.Equal(3, refreshed.Count);
        Assert.Equal(3, refreshed["Extra"]);
        Assert.Equal(updated, worksheet.OuterXml);
    }

    private sealed class ImplicitRowRecord {
        public decimal? Id { get; set; }
        public string? Name { get; set; }
        public decimal? Amount { get; set; }
    }
}
