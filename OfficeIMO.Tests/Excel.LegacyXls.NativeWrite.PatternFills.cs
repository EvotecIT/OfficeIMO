using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesGray125PatternFillWhenColorsAreSpecified() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("PatternFills");
                    sheet.CellValue(1, 1, "Gray125");
                    sheet.CellAt(1, 1).SetFillColor("#EEEEEE");

                    Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    CellFormat baseFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)baseStyleIndex);
                    var gray125Fill = new Fill(new PatternFill {
                        PatternType = PatternValues.Gray125,
                        ForegroundColor = new ForegroundColor { Rgb = "FF123456" },
                        BackgroundColor = new BackgroundColor { Rgb = "FFABCDEF" }
                    });

                    stylesheet.Fills!.Append(gray125Fill);
                    stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();
                    var gray125Format = (CellFormat)baseFormat.CloneNode(true);
                    gray125Format.FillId = stylesheet.Fills.Count!.Value - 1;
                    gray125Format.ApplyFill = true;
                    stylesheet.CellFormats.Append(gray125Format);
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
                LegacyXlsCell filledCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[filledCell.StyleIndex];
                Assert.True(cellFormat.ApplyFill);
                Assert.Equal((byte)17, cellFormat.FillPattern);
                Assert.True(workbook.TryResolveColor(cellFormat.FillForegroundColorIndex, out string? foregroundColor));
                Assert.Equal("FF123456", foregroundColor);
                Assert.True(workbook.TryResolveColor(cellFormat.FillBackgroundColorIndex, out string? backgroundColor));
                Assert.Equal("FFABCDEF", backgroundColor);

                Cell projectedCell = result.Document.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                Stylesheet projectedStylesheet = result.Document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                CellFormat projectedFormat = projectedStylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)projectedCell.StyleIndex!.Value);
                Fill projectedFill = projectedStylesheet.Fills!.Elements<Fill>().ElementAt((int)projectedFormat.FillId!.Value);
                Assert.Equal(PatternValues.Gray125, projectedFill.PatternFill!.PatternType!.Value);
                Assert.Equal("FF123456", projectedFill.PatternFill.ForegroundColor!.Rgb!.Value);
                Assert.Equal("FFABCDEF", projectedFill.PatternFill.BackgroundColor!.Rgb!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
