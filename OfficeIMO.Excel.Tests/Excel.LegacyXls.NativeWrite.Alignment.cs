using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesExplicitGeneralHorizontalAlignment() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("Alignment");
                    sheet.CellValue(1, 1, "General horizontal");
                    sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                    Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    CellFormat baseFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)baseStyleIndex);
                    var generalFormat = (CellFormat)baseFormat.CloneNode(true);
                    generalFormat.Alignment = new Alignment {
                        Horizontal = HorizontalAlignmentValues.General,
                        Vertical = VerticalAlignmentValues.Center,
                        WrapText = true
                    };
                    generalFormat.ApplyAlignment = true;
                    stylesheet.CellFormats.Append(generalFormat);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell alignedCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[alignedCell.StyleIndex];
                Assert.True(cellFormat.ApplyAlignment);
                Assert.Equal((byte)0, cellFormat.HorizontalAlignment);
                Assert.Equal((byte)1, cellFormat.VerticalAlignment);
                Assert.True(cellFormat.WrapText);

                Cell projectedCell = result.Document.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                Stylesheet projectedStylesheet = result.Document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                CellFormat projectedFormat = projectedStylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)projectedCell.StyleIndex!.Value);
                Assert.True(projectedFormat.ApplyAlignment!.Value);
                Assert.NotNull(projectedFormat.Alignment);
                Assert.Null(projectedFormat.Alignment!.Horizontal);
                Assert.Equal(VerticalAlignmentValues.Center, projectedFormat.Alignment.Vertical!.Value);
                Assert.True(projectedFormat.Alignment.WrapText!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
