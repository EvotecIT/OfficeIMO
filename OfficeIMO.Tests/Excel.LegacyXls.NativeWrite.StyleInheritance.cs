using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_InheritsSparseCellFormatFromParentStyleFormat() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("StyleParent");
                    sheet.CellValue(1, 1, 45291d);

                    WorkbookStylesPart stylesPart = document.WorkbookPartRoot!.WorkbookStylesPart ?? document.WorkbookPartRoot.AddNewPart<WorkbookStylesPart>();
                    Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                    EnsureLegacyXlsNativeWriteStyleCollections(stylesheet);

                    stylesheet.Fonts!.Append(new Font(new Bold(), new FontName { Val = "Arial" }));
                    stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();
                    uint parentFontId = stylesheet.Fonts.Count!.Value - 1;

                    stylesheet.Fills!.Append(new Fill(new PatternFill {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor { Rgb = "FFABCDEF" },
                        BackgroundColor = new BackgroundColor { Indexed = 64U }
                    }));
                    stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();
                    uint parentFillId = stylesheet.Fills.Count!.Value - 1;

                    stylesheet.Borders!.Append(new Border(
                        new LeftBorder(new Color { Rgb = "FF336699" }) { Style = BorderStyleValues.Thin },
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder()));
                    stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();
                    uint parentBorderId = stylesheet.Borders.Count!.Value - 1;

                    uint parentFormatId = stylesheet.CellStyleFormats!.Count?.Value ?? (uint)stylesheet.CellStyleFormats.Count();
                    stylesheet.CellStyleFormats.Append(new CellFormat {
                        NumberFormatId = 14U,
                        FontId = parentFontId,
                        FillId = parentFillId,
                        BorderId = parentBorderId,
                        ApplyNumberFormat = true,
                        ApplyFont = true,
                        ApplyFill = true,
                        ApplyBorder = true,
                        ApplyAlignment = true,
                        ApplyProtection = true,
                        QuotePrefix = true,
                        Alignment = new Alignment {
                            Horizontal = HorizontalAlignmentValues.Center,
                            WrapText = true
                        },
                        Protection = new Protection {
                            Locked = false,
                            Hidden = true
                        }
                    });
                    stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();

                    stylesheet.CellFormats!.Append(new CellFormat {
                        FormatId = parentFormatId
                    });
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    uint childFormatIndex = stylesheet.CellFormats.Count!.Value - 1;

                    Cell openXmlCell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                    openXmlCell.StyleIndex = childFormatIndex;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell cell = Assert.Single(worksheet.Cells, item => item.Row == 1 && item.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[cell.StyleIndex];

                Assert.Equal((ushort)14, cellFormat.NumberFormatId);
                Assert.True(cellFormat.ApplyNumberFormat);
                Assert.True(cellFormat.ApplyFont);
                Assert.True(GetLegacyFont(workbook, cellFormat.FontIndex).Bold);
                Assert.True(cellFormat.ApplyFill);
                Assert.Equal((byte)1, cellFormat.FillPattern);
                Assert.True(workbook.TryResolveColor(cellFormat.FillForegroundColorIndex, out string? fillColor));
                Assert.Equal("FFABCDEF", fillColor);
                Assert.True(cellFormat.ApplyBorder);
                Assert.NotNull(cellFormat.Border);
                Assert.True(workbook.TryResolveColor(cellFormat.Border!.LeftColorIndex, out string? borderColor));
                Assert.Equal("FF336699", borderColor);
                Assert.True(cellFormat.ApplyAlignment);
                Assert.Equal((byte)2, cellFormat.HorizontalAlignment);
                Assert.True(cellFormat.WrapText);
                Assert.True(cellFormat.ApplyProtection);
                Assert.False(cellFormat.Locked);
                Assert.True(cellFormat.FormulaHidden);
                Assert.True(cellFormat.QuotePrefix);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesStyledBlankCells() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("StyledBlank");
                    sheet.CellValue(2, 2, "temporary");

                    WorkbookStylesPart stylesPart = document.WorkbookPartRoot!.WorkbookStylesPart ?? document.WorkbookPartRoot.AddNewPart<WorkbookStylesPart>();
                    Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                    EnsureLegacyXlsNativeWriteStyleCollections(stylesheet);

                    stylesheet.Fills!.Append(new Fill(new PatternFill {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor { Rgb = "FF00AA88" },
                        BackgroundColor = new BackgroundColor { Indexed = 64U }
                    }));
                    stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();
                    uint fillId = stylesheet.Fills.Count!.Value - 1;

                    stylesheet.CellFormats!.Append(new CellFormat {
                        FillId = fillId,
                        ApplyFill = true
                    });
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    uint blankStyleIndex = stylesheet.CellFormats.Count!.Value - 1;

                    Cell blankCell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>()
                        .Single(cell => string.Equals(cell.CellReference?.Value, "B2", StringComparison.OrdinalIgnoreCase));
                    blankCell.StyleIndex = blankStyleIndex;
                    blankCell.DataType = null;
                    blankCell.RemoveAllChildren<CellValue>();
                    blankCell.RemoveAllChildren<InlineString>();
                    blankCell.RemoveAllChildren<CellFormula>();
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell blank = Assert.Single(worksheet.Cells, cell => cell.Row == 2 && cell.Column == 2);
                Assert.Equal(LegacyXlsCellValueKind.Blank, blank.Kind);
                Assert.Null(blank.Value);

                LegacyXlsCellFormat cellFormat = result.Workbook.CellFormats[blank.StyleIndex];
                Assert.True(cellFormat.ApplyFill);
                Assert.Equal((byte)1, cellFormat.FillPattern);
                Assert.True(result.Workbook.TryResolveColor(cellFormat.FillForegroundColorIndex, out string? fillColor));
                Assert.Equal("FF00AA88", fillColor);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesStyledEmptyStringCellsAsBlanks() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("EmptyString");
                    sheet.CellValue(2, 2, "temporary");

                    WorkbookStylesPart stylesPart = document.WorkbookPartRoot!.WorkbookStylesPart ?? document.WorkbookPartRoot.AddNewPart<WorkbookStylesPart>();
                    Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                    EnsureLegacyXlsNativeWriteStyleCollections(stylesheet);

                    stylesheet.Fills!.Append(new Fill(new PatternFill {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor { Rgb = "FFAA5500" },
                        BackgroundColor = new BackgroundColor { Indexed = 64U }
                    }));
                    stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();
                    uint fillId = stylesheet.Fills.Count!.Value - 1;

                    stylesheet.CellFormats!.Append(new CellFormat {
                        FillId = fillId,
                        ApplyFill = true
                    });
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    uint blankStyleIndex = stylesheet.CellFormats.Count!.Value - 1;

                    Cell blankCell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>()
                        .Single(cell => string.Equals(cell.CellReference?.Value, "B2", StringComparison.OrdinalIgnoreCase));
                    blankCell.StyleIndex = blankStyleIndex;
                    blankCell.DataType = CellValues.InlineString;
                    blankCell.RemoveAllChildren<CellValue>();
                    blankCell.RemoveAllChildren<InlineString>();
                    blankCell.RemoveAllChildren<CellFormula>();
                    blankCell.Append(new InlineString(new Text(string.Empty)));
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell blank = Assert.Single(worksheet.Cells, cell => cell.Row == 2 && cell.Column == 2);
                Assert.Equal(LegacyXlsCellValueKind.Blank, blank.Kind);
                Assert.Null(blank.Value);

                LegacyXlsCellFormat cellFormat = result.Workbook.CellFormats[blank.StyleIndex];
                Assert.True(cellFormat.ApplyFill);
                Assert.Equal((byte)1, cellFormat.FillPattern);
                Assert.True(result.Workbook.TryResolveColor(cellFormat.FillForegroundColorIndex, out string? fillColor));
                Assert.Equal("FFAA5500", fillColor);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        private static void EnsureLegacyXlsNativeWriteStyleCollections(Stylesheet stylesheet) {
            stylesheet.Fonts ??= new Fonts(new Font());
            stylesheet.Fills ??= new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }));
            stylesheet.Borders ??= new Borders(new Border());
            stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat());
            stylesheet.CellFormats ??= new CellFormats(new CellFormat());

            if (!stylesheet.CellStyleFormats.Elements<CellFormat>().Any()) {
                stylesheet.CellStyleFormats.Append(new CellFormat());
            }
        }
    }
}
