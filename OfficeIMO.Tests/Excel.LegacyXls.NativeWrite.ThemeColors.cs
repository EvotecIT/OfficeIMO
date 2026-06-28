using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesThemeOrTintFontFillAndBorderColors() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("ThemeStyles");
                    sheet.CellValue(1, 1, "Theme styled");
                    sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                    Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    CellFormat baseFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)baseStyleIndex);

                    var themedFont = new Font(new Color { Theme = 4U, Tint = -0.4D });
                    stylesheet.Fonts!.Append(themedFont);
                    stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

                    var themedFill = new Fill(new PatternFill {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor { Theme = 1U, Tint = 0.4D },
                        BackgroundColor = new BackgroundColor { Indexed = 64U }
                    });
                    stylesheet.Fills!.Append(themedFill);
                    stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

                    var themedBorder = new Border(
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(new Color { Theme = 5U, Tint = 0.2D }) { Style = BorderStyleValues.Thin },
                        new BottomBorder());
                    stylesheet.Borders!.Append(themedBorder);
                    stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();

                    var themedFormat = (CellFormat)baseFormat.CloneNode(true);
                    themedFormat.FontId = stylesheet.Fonts.Count!.Value - 1;
                    themedFormat.FillId = stylesheet.Fills.Count!.Value - 1;
                    themedFormat.BorderId = stylesheet.Borders.Count!.Value - 1;
                    themedFormat.ApplyFont = true;
                    themedFormat.ApplyFill = true;
                    themedFormat.ApplyBorder = true;
                    stylesheet.CellFormats.Append(themedFormat);
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
                LegacyXlsCell themedCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[themedCell.StyleIndex];

                LegacyXlsFont font = GetLegacyFont(workbook, cellFormat.FontIndex);
                Assert.True(workbook.TryResolveColor(font.ColorIndex, out string? fontColor));
                Assert.Equal("FF2F4D71", fontColor);

                Assert.True(cellFormat.ApplyFill);
                Assert.Equal((byte)1, cellFormat.FillPattern);
                Assert.True(workbook.TryResolveColor(cellFormat.FillForegroundColorIndex, out string? fillColor));
                Assert.Equal("FF666666", fillColor);

                Assert.True(cellFormat.ApplyBorder);
                Assert.NotNull(cellFormat.Border);
                Assert.True(workbook.TryResolveColor(cellFormat.Border!.TopColorIndex, out string? topBorderColor));
                Assert.Equal("FFCD7371", topBorderColor);

                ExcelCellStyleSnapshot projectedStyle = result.Document.Sheets[0].GetCellStyle(1, 1);
                Assert.Equal("FF2F4D71", projectedStyle.FontColorArgb);
                Assert.Equal("FF666666", projectedStyle.FillColorArgb);
                Assert.NotNull(projectedStyle.Border);
                Assert.Equal("FFCD7371", projectedStyle.Border!.Top!.ColorArgb);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
