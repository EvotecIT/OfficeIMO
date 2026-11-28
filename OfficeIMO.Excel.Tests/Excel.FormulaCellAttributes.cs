using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelFormulaCellAttributes {
        [Fact]
        public void RecalculationReplacesAnUnreadableCalculateFlag() {
            byte[] bytes = Workbook(sheet => sheet.Cell(1, 1).SetFormula("1+1"), cell => {
                cell.CellFormula!.CalculateCell = new BooleanValue { InnerText = "invalid" };
            });

            using var document = ExcelDocument.Load(new MemoryStream(bytes, false));
            Assert.Equal(1, document.Sheets[0].RecalculateSupportedFormulas());

            Assert.True(document.Sheets[0].TryGetCachedFormulaValue(1, 1, out string? value));
            Assert.Equal("2", value);
            CellFormula formula = document.Sheets[0].WorksheetPart.Worksheet!.Descendants<CellFormula>().Single();
            Assert.False(formula.CalculateCell!.Value);
        }

        [Fact]
        public void AnArrayFormulaWithAnOutOfRangeMetadataIndexStillCalculates() {
            byte[] bytes = Workbook(sheet => {
                sheet.Cell(1, 1).SetValue(2d);
                sheet.Cell(2, 1).SetValue(3d);
                sheet.SetArrayFormula("B1:B2", "A1:A2");
            }, cell => {
                if (cell.CellReference?.Value == "B1") cell.CellMetaIndex = new UInt32Value { InnerText = "4294967296" };
            });

            using var document = ExcelDocument.Load(new MemoryStream(bytes, false));
            document.Sheets[0].Cell(2, 1).SetValue(7d);
            document.Sheets[0].RecalculateSupportedFormulas();

            Assert.Equal(7d, document.Sheets[0].Cell(2, 2).GetValue().Value);
        }

        private static byte[] Workbook(Action<ExcelSheet> write, Action<Cell> damage) {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                write(document.AddSheet("Data"));
                document.Save();
            }
            using (var package = SpreadsheetDocument.Open(stream, true)) {
                foreach (Cell cell in package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToList()) damage(cell);
            }
            return stream.ToArray();
        }
    }
}
