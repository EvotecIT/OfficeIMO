using System.Data;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace OfficeIMO.Excel.Tests {
    public partial class ExcelTests {
        [Theory]
        [InlineData(false, 1)]
        [InlineData(true, 1)]
        [InlineData(false, 512)]
        public void DirectPackageWriter_EscapedClusterThenLongPlainTextPreservesLaterEscapes(bool compact, int rowCount) {
            string source = "A&<tag>" + new string('q', 2048)
                + "\r&<later>" + new string('z', 2048) + '\u0001' + "tail";
            string expected = source.Replace("\u0001", string.Empty);
            var table = new DataTable();
            table.Columns.Add("Notes", typeof(string));
            for (int row = 0; row < rowCount; row++) table.Rows.Add(source);

            using var output = new MemoryStream();
            using (var reader = table.CreateDataReader()) {
                ExcelDocument.WriteDataReader(output, reader, compact
                    ? new ExcelTabularWriteOptions { IncludeCellReferences = false, UseSharedStrings = false }
                    : null);
            }

            output.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            WorksheetPart sheet = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell[] cells = sheet.Worksheet.Descendants<Cell>().ToArray();
            Assert.Equal(rowCount + 1, cells.Length);
            Cell cell = cells[1];
            if (!compact) Assert.Equal("A2", cell.CellReference?.Value);
            if (rowCount > 1) {
                Assert.Equal(CellValues.SharedString, cell.DataType?.Value);
                Assert.NotNull(spreadsheet.WorkbookPart.SharedStringTablePart);
            }
            string actual = cell.DataType?.Value == CellValues.SharedString
                ? spreadsheet.WorkbookPart.SharedStringTablePart!.SharedStringTable!
                    .Elements<SharedStringItem>()
                    .ElementAt(int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture)).InnerText
                : cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? string.Empty;
            Assert.Equal(expected, actual);
            using var xmlReader = new StreamReader(rowCount > 1
                ? spreadsheet.WorkbookPart.SharedStringTablePart!.GetStream()
                : sheet.GetStream());
            string textXml = xmlReader.ReadToEnd();
            Assert.Contains("&#xD;", textXml, StringComparison.Ordinal);
            Assert.DoesNotContain('\u0001', textXml);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void DirectPackageWriter_PreservesUnicodeAndBlankValuesInTypedRows() {
            using var output = new MemoryStream();
            string emoji = char.ConvertFromUtf32(0x1F680);
            var rows = new[] {
                new DirectPackageWriterRow(1, "Zażółć", "東京", new DateTime(2026, 7, 10), 12.5, 2, true, "A&B < " + emoji + "\r\nnext\rreturn"),
                new DirectPackageWriterRow(2, null, "München", new DateTime(2026, 7, 11), 20.75, 3, false, "Plain")
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Id", row => row.Id),
                    ("Region", row => row.Region),
                    ("Owner", row => row.Owner),
                    ("CreatedOn", row => row.CreatedOn),
                    ("Amount", row => row.Amount),
                    ("Units", row => row.Units),
                    ("Active", row => row.Active),
                    ("Notes", row => row.Notes));

                document.Save(output);
                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            output.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet
                .Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("Zażółć", GetDirectCellText(cells["B2"]));
            Assert.Equal("東京", GetDirectCellText(cells["C2"]));
            Assert.Equal("A&B < " + emoji + "\r\nnext\rreturn", GetDirectCellText(cells["H2"]));
            Assert.False(cells.ContainsKey("B3"));
            Assert.Equal("München", GetDirectCellText(cells["C3"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(output.ToArray());
            Assert.True(reader.Read());
            Assert.Equal("Zażółć", reader.GetString(1));
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(1));
            Assert.Equal("München", reader.GetString(2));
            Assert.False(reader.Read());

            output.Position = 0;
            using var reopened = ExcelDocument.Load(output);
            Assert.True(reopened["Data"].TryGetCellText(2, 8, out string notes));
            Assert.Equal("A&B < " + emoji + "\r\nnext\rreturn", notes);
        }

        private static string? GetDirectCellText(Cell cell)
            => cell.InlineString?.InnerText ?? cell.CellValue?.Text;

        private sealed record DirectPackageWriterRow(
            int Id,
            string? Region,
            string Owner,
            DateTime CreatedOn,
            double Amount,
            int Units,
            bool Active,
            string Notes);
    }
}
