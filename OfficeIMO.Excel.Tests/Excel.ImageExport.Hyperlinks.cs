using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelImageExportHyperlinkTests {
        [Fact]
        public void ExcelRange_ImageExportAddsHyperlinkVisualHintWhenCellStyleDoesNot() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Links");
            sheet.SetColumnWidth(1, 20);
            sheet.SetHyperlink(1, 1, "https://example.org/spec", display: "Spec", style: false);

            string hintedSvg = sheet.Range("A1:A1").ToSvg(new ExcelImageExportOptions { ShowGridlines = false });
            string plainSvg = sheet.Range("A1:A1").ToSvg(new ExcelImageExportOptions {
                ShowGridlines = false,
                ShowHyperlinkHints = false
            });
            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A1").CreateVisualSnapshot();

            Assert.NotNull(snapshot.Cells[0].Hyperlink);
            Assert.True(snapshot.Cells[0].Hyperlink!.IsExternal);
            Assert.Equal("https://example.org/spec", snapshot.Cells[0].Hyperlink!.Target);
            Assert.Contains("#0563C1", hintedSvg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("text-decoration=\"underline\"", hintedSvg, StringComparison.Ordinal);
            Assert.DoesNotContain("#0563C1", plainSvg, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("text-decoration=\"underline\"", plainSvg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelInspectionSnapshotExpandsWorksheetHyperlinkRangesToCells() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Links");
                sheet.CellValue(1, 1, "First");
                sheet.CellValue(1, 2, "Second");
                sheet.SetHyperlink(1, 1, "https://example.org/range", display: "First", style: false);
                document.Save(false);
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                Hyperlink hyperlink = worksheetPart.Worksheet!.Descendants<Hyperlink>().Single();
                hyperlink.Reference = "A1:B1";
                worksheetPart.Worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                ExcelWorksheetSnapshot worksheet = snapshot.Worksheets.Single();
                ExcelCellSnapshot first = worksheet.Cells.Single(cell => cell.Row == 1 && cell.Column == 1);
                ExcelCellSnapshot second = worksheet.Cells.Single(cell => cell.Row == 1 && cell.Column == 2);

                Assert.NotNull(first.Hyperlink);
                Assert.NotNull(second.Hyperlink);
                Assert.Equal("https://example.org/range", first.Hyperlink!.Target);
                Assert.Equal("https://example.org/range", second.Hyperlink!.Target);
            }
        }
    }
}
