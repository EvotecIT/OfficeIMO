using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesDivergentWorkbookViews() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    document.AddWorksheet("First").CellValue(1, 1, "First");
                    document.AddWorksheet("Second").CellValue(1, 1, "Second");

                    Workbook workbook = document.WorkbookRoot;
                    workbook.RemoveAllChildren<BookViews>();
                    var bookViews = new BookViews(
                        new WorkbookView {
                            ActiveTab = 1U,
                            FirstSheet = 1U,
                            XWindow = 12,
                            YWindow = 34,
                            WindowWidth = 6000U,
                            WindowHeight = 4000U,
                            ShowHorizontalScroll = false,
                            ShowVerticalScroll = true,
                            ShowSheetTabs = true,
                            TabRatio = 650U
                        },
                        new WorkbookView {
                            ActiveTab = 0U,
                            FirstSheet = 0U,
                            XWindow = 56,
                            YWindow = 78,
                            WindowWidth = 7000U,
                            WindowHeight = 4500U,
                            Visibility = VisibilityValues.Hidden,
                            Minimized = true,
                            ShowHorizontalScroll = true,
                            ShowVerticalScroll = false,
                            ShowSheetTabs = false,
                            TabRatio = 500U
                        });
                    Sheets? sheets = workbook.GetFirstChild<Sheets>();
                    if (sheets != null) {
                        workbook.InsertBefore(bookViews, sheets);
                    } else {
                        workbook.Append(bookViews);
                    }

                    workbook.Save();
                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                Assert.Equal(2, result.Workbook.Windows.Count);
                LegacyXlsWorkbookWindow firstWindow = result.Workbook.Windows[0];
                Assert.Equal((ushort)1, firstWindow.ActiveSheetIndex);
                Assert.Equal((ushort)1, firstWindow.FirstVisibleSheetTabIndex);
                Assert.Equal((short)12, firstWindow.HorizontalPositionTwips);
                Assert.Equal((short)34, firstWindow.VerticalPositionTwips);
                Assert.Equal((short)6000, firstWindow.WidthTwips);
                Assert.Equal((short)4000, firstWindow.HeightTwips);
                Assert.False(firstWindow.HorizontalScrollBarVisible);
                Assert.True(firstWindow.VerticalScrollBarVisible);
                Assert.True(firstWindow.SheetTabsVisible);
                Assert.Equal((ushort)650, firstWindow.SheetTabRatio);

                LegacyXlsWorkbookWindow secondWindow = result.Workbook.Windows[1];
                Assert.Equal((ushort)0, secondWindow.ActiveSheetIndex);
                Assert.Equal((ushort)0, secondWindow.FirstVisibleSheetTabIndex);
                Assert.Equal((short)56, secondWindow.HorizontalPositionTwips);
                Assert.Equal((short)78, secondWindow.VerticalPositionTwips);
                Assert.Equal((short)7000, secondWindow.WidthTwips);
                Assert.Equal((short)4500, secondWindow.HeightTwips);
                Assert.True(secondWindow.Hidden);
                Assert.True(secondWindow.Minimized);
                Assert.True(secondWindow.HorizontalScrollBarVisible);
                Assert.False(secondWindow.VerticalScrollBarVisible);
                Assert.False(secondWindow.SheetTabsVisible);
                Assert.Equal((ushort)500, secondWindow.SheetTabRatio);

                WorkbookView[] projectedViews = result.Document.WorkbookRoot
                    .GetFirstChild<BookViews>()!
                    .Elements<WorkbookView>()
                    .ToArray();
                Assert.Equal(2, projectedViews.Length);
                Assert.Equal(1U, projectedViews[0].ActiveTab!.Value);
                Assert.Equal(1U, projectedViews[0].FirstSheet!.Value);
                Assert.Equal(650U, projectedViews[0].TabRatio!.Value);
                Assert.Equal(0U, projectedViews[1].ActiveTab!.Value);
                Assert.Equal(0U, projectedViews[1].FirstSheet!.Value);
                Assert.Equal(VisibilityValues.Hidden, projectedViews[1].Visibility!.Value);
                Assert.True(projectedViews[1].Minimized!.Value);
                Assert.False(projectedViews[1].ShowVerticalScroll!.Value);
                Assert.False(projectedViews[1].ShowSheetTabs!.Value);
                Assert.Equal(500U, projectedViews[1].TabRatio!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_CollapsesEquivalentWorksheetViews() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Views");
                    sheet.CellValue(1, 1, "Equivalent worksheet views");

                    Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                    worksheet.RemoveAllChildren<SheetViews>();
                    var sheetViews = new SheetViews(
                        new SheetView {
                            WorkbookViewId = 0U,
                            ShowGridLines = false,
                            ShowRowColHeaders = false,
                            ShowZeros = false,
                            RightToLeft = true,
                            DefaultGridColor = false,
                            ColorId = 22U,
                            ShowOutlineSymbols = false,
                            TabSelected = true,
                            View = SheetViewValues.PageBreakPreview,
                            ZoomScale = 125U,
                            ZoomScaleNormal = 90U,
                            TopLeftCell = "D6"
                        },
                        new SheetView {
                            WorkbookViewId = 1U,
                            ShowGridLines = false,
                            ShowRowColHeaders = false,
                            ShowZeros = false,
                            RightToLeft = true,
                            DefaultGridColor = false,
                            ColorId = 22U,
                            ShowOutlineSymbols = false,
                            TabSelected = true,
                            View = SheetViewValues.PageBreakPreview,
                            ZoomScale = 125U,
                            ZoomScaleNormal = 90U,
                            TopLeftCell = "D6"
                        });
                    SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        worksheet.InsertBefore(sheetViews, sheetData);
                    } else {
                        worksheet.InsertAt(sheetViews, 0);
                    }

                    worksheet.Save();
                    document.Save(xlsOutputPath);
                }

                using ExcelDocument loaded = ExcelDocument.Load(xlsOutputPath);
                ExcelSheet loadedSheet = loaded.Sheets.Single();

                Assert.True(loaded.SourceFormat == ExcelFileFormat.Xls);
                ExcelWorksheetViewInfo view = loadedSheet.GetViewInfo();
                Assert.False(view.ShowGridlines);
                Assert.False(loadedSheet.RowColumnHeadingsVisible);
                Assert.False(loadedSheet.ZeroValuesVisible);
                Assert.True(view.RightToLeft);
                Assert.Equal(125U, view.ZoomScale);
                Assert.Equal(90U, view.ZoomScaleNormal);
                Assert.Equal("pageBreakPreview", view.View);

                SheetView loadedSheetView = loadedSheet.WorksheetPart.Worksheet
                    .GetFirstChild<SheetViews>()!
                    .GetFirstChild<SheetView>()!;
                Assert.False(loadedSheetView.DefaultGridColor!.Value);
                Assert.Equal(22U, loadedSheetView.ColorId!.Value);
                Assert.False(loadedSheetView.ShowOutlineSymbols!.Value);
                Assert.True(loadedSheetView.TabSelected!.Value);
                Assert.Equal("D6", loadedSheetView.TopLeftCell!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesDivergentWorksheetViews() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Views");
                    sheet.CellValue(1, 1, "Divergent worksheet views");

                    Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                    worksheet.RemoveAllChildren<SheetViews>();
                    var sheetViews = new SheetViews(
                        new SheetView {
                            WorkbookViewId = 0U,
                            ShowFormulas = true,
                            ShowGridLines = false,
                            ShowRowColHeaders = false,
                            ShowZeros = false,
                            RightToLeft = true,
                            DefaultGridColor = false,
                            ColorId = 22U,
                            ShowOutlineSymbols = false,
                            TabSelected = true,
                            View = SheetViewValues.Normal,
                            ZoomScaleNormal = 90U,
                            TopLeftCell = "D6"
                        },
                        new SheetView {
                            WorkbookViewId = 1U,
                            ShowFormulas = false,
                            ShowGridLines = true,
                            ShowRowColHeaders = true,
                            ShowZeros = true,
                            RightToLeft = false,
                            DefaultGridColor = true,
                            ShowOutlineSymbols = true,
                            TabSelected = false,
                            View = SheetViewValues.PageBreakPreview,
                            ZoomScale = 125U,
                            ZoomScaleNormal = 110U,
                            TopLeftCell = "B3"
                        },
                        new SheetView {
                            WorkbookViewId = 2U,
                            ShowFormulas = false,
                            ShowGridLines = false,
                            ShowRowColHeaders = true,
                            ShowZeros = true,
                            RightToLeft = false,
                            DefaultGridColor = true,
                            ShowOutlineSymbols = false,
                            TabSelected = false,
                            View = SheetViewValues.Normal,
                            ZoomScale = 140U,
                            TopLeftCell = "E7"
                        });
                    SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        worksheet.InsertBefore(sheetViews, sheetData);
                    } else {
                        worksheet.InsertAt(sheetViews, 0);
                    }

                    worksheet.Save();
                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = result.Workbook.Worksheets.Single();
                Assert.Equal(3, legacySheet.WindowViews.Count);
                Assert.True(legacySheet.WindowViews[0].ShowFormulas);
                Assert.False(legacySheet.WindowViews[0].ShowGridLines);
                Assert.False(legacySheet.WindowViews[0].ShowRowColumnHeadings);
                Assert.False(legacySheet.WindowViews[0].ShowZeroValues);
                Assert.True(legacySheet.WindowViews[0].RightToLeft);
                Assert.False(legacySheet.WindowViews[0].DefaultGridColor);
                Assert.Equal((ushort)22, legacySheet.WindowViews[0].GridLineColorIndex);
                Assert.False(legacySheet.WindowViews[0].ShowOutlineSymbols);
                Assert.True(legacySheet.WindowViews[0].TabSelected);
                Assert.False(legacySheet.WindowViews[0].PageBreakPreview);
                Assert.Equal(5, legacySheet.WindowViews[0].FirstVisibleRow);
                Assert.Equal(3, legacySheet.WindowViews[0].FirstVisibleColumn);
                Assert.Equal(90U, legacySheet.WindowViews[0].ZoomScaleNormal);

                Assert.False(legacySheet.WindowViews[1].ShowFormulas);
                Assert.True(legacySheet.WindowViews[1].ShowGridLines);
                Assert.True(legacySheet.WindowViews[1].ShowRowColumnHeadings);
                Assert.True(legacySheet.WindowViews[1].ShowZeroValues);
                Assert.False(legacySheet.WindowViews[1].RightToLeft);
                Assert.True(legacySheet.WindowViews[1].DefaultGridColor);
                Assert.True(legacySheet.WindowViews[1].ShowOutlineSymbols);
                Assert.False(legacySheet.WindowViews[1].TabSelected);
                Assert.True(legacySheet.WindowViews[1].PageBreakPreview);
                Assert.Equal(2, legacySheet.WindowViews[1].FirstVisibleRow);
                Assert.Equal(1, legacySheet.WindowViews[1].FirstVisibleColumn);
                Assert.Equal(125U, legacySheet.WindowViews[1].ZoomScale);
                Assert.Equal(110U, legacySheet.WindowViews[1].ZoomScaleNormal);

                Assert.False(legacySheet.WindowViews[2].ShowFormulas);
                Assert.False(legacySheet.WindowViews[2].ShowGridLines);
                Assert.True(legacySheet.WindowViews[2].ShowRowColumnHeadings);
                Assert.True(legacySheet.WindowViews[2].ShowZeroValues);
                Assert.False(legacySheet.WindowViews[2].RightToLeft);
                Assert.True(legacySheet.WindowViews[2].DefaultGridColor);
                Assert.False(legacySheet.WindowViews[2].ShowOutlineSymbols);
                Assert.False(legacySheet.WindowViews[2].TabSelected);
                Assert.False(legacySheet.WindowViews[2].PageBreakPreview);
                Assert.Equal(6, legacySheet.WindowViews[2].FirstVisibleRow);
                Assert.Equal(4, legacySheet.WindowViews[2].FirstVisibleColumn);
                Assert.Null(legacySheet.WindowViews[2].ZoomScale);
                Assert.Equal(140U, legacySheet.WindowViews[2].ZoomScaleNormal);

                SheetView[] projectedViews = result.Document.Sheets.Single().WorksheetPart.Worksheet
                    .GetFirstChild<SheetViews>()!
                    .Elements<SheetView>()
                    .ToArray();
                Assert.Equal(3, projectedViews.Length);
                Assert.True(projectedViews[0].ShowFormulas!.Value);
                Assert.False(projectedViews[0].ShowGridLines!.Value);
                Assert.False(projectedViews[0].ShowRowColHeaders!.Value);
                Assert.False(projectedViews[0].ShowZeros!.Value);
                Assert.True(projectedViews[0].RightToLeft!.Value);
                Assert.False(projectedViews[0].DefaultGridColor!.Value);
                Assert.Equal(22U, projectedViews[0].ColorId!.Value);
                Assert.False(projectedViews[0].ShowOutlineSymbols!.Value);
                Assert.True(projectedViews[0].TabSelected!.Value);
                Assert.Equal(SheetViewValues.Normal, projectedViews[0].View!.Value);
                Assert.Equal("D6", projectedViews[0].TopLeftCell!.Value);
                Assert.Equal(90U, projectedViews[0].ZoomScaleNormal!.Value);

                Assert.False(projectedViews[1].ShowFormulas!.Value);
                Assert.True(projectedViews[1].ShowGridLines!.Value);
                Assert.True(projectedViews[1].ShowRowColHeaders!.Value);
                Assert.True(projectedViews[1].ShowZeros!.Value);
                Assert.False(projectedViews[1].RightToLeft!.Value);
                Assert.True(projectedViews[1].DefaultGridColor!.Value);
                Assert.True(projectedViews[1].ShowOutlineSymbols!.Value);
                Assert.False(projectedViews[1].TabSelected!.Value);
                Assert.Equal(SheetViewValues.PageBreakPreview, projectedViews[1].View!.Value);
                Assert.Equal("B3", projectedViews[1].TopLeftCell!.Value);
                Assert.Equal(125U, projectedViews[1].ZoomScale!.Value);
                Assert.Equal(110U, projectedViews[1].ZoomScaleNormal!.Value);

                Assert.False(projectedViews[2].ShowFormulas!.Value);
                Assert.False(projectedViews[2].ShowGridLines!.Value);
                Assert.True(projectedViews[2].ShowRowColHeaders!.Value);
                Assert.True(projectedViews[2].ShowZeros!.Value);
                Assert.False(projectedViews[2].RightToLeft!.Value);
                Assert.True(projectedViews[2].DefaultGridColor!.Value);
                Assert.False(projectedViews[2].ShowOutlineSymbols!.Value);
                Assert.False(projectedViews[2].TabSelected!.Value);
                Assert.Equal(SheetViewValues.Normal, projectedViews[2].View!.Value);
                Assert.Equal("E7", projectedViews[2].TopLeftCell!.Value);
                Assert.Equal(140U, projectedViews[2].ZoomScale!.Value);
                Assert.Equal(140U, projectedViews[2].ZoomScaleNormal!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
