using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace OfficeIMO.Excel.Tests {
    public partial class ExcelTests {
        [Fact]
        public void DirectPackageWriter_PreservesUnicodeAndBlankValuesInTypedRows() {
            using var output = new MemoryStream();
            string emoji = char.ConvertFromUtf32(0x1F680);
            var rows = new[] {
                new DirectPackageWriterRow(1, "Zażółć", "東京", new DateTime(2026, 7, 10), 12.5, 2, true, "A&B < " + emoji),
                new DirectPackageWriterRow(2, null, "München", new DateTime(2026, 7, 11), 20.75, 3, false, "Plain")
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorkSheet("Data");
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
            Assert.Equal("A&B < " + emoji, GetDirectCellText(cells["H2"]));
            Assert.Equal(string.Empty, GetDirectCellText(cells["B3"]));
            Assert.Equal("München", GetDirectCellText(cells["C3"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
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
