using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal void CleanupWorkbookViewArtifacts(bool save = true) {
            var workbook = _workBookPart.Workbook;
            bool hasSheetViews = _workBookPart.WorksheetParts.Any(worksheetPart =>
                worksheetPart.Worksheet.GetFirstChild<SheetViews>()?.Elements<SheetView>().Any() == true);

            var workbookViews = workbook.GetFirstChild<BookViews>();
            bool workbookChanged = false;

            if (hasSheetViews) {
                if (workbookViews == null) {
                    workbookViews = new BookViews();
                    var sheets = workbook.GetFirstChild<Sheets>();
                    if (sheets != null) {
                        workbook.InsertBefore(workbookViews, sheets);
                    } else {
                        workbook.Append(workbookViews);
                    }
                    workbookChanged = true;
                }

                if (!workbookViews.Elements<WorkbookView>().Any()) {
                    workbookViews.Append(new WorkbookView());
                    workbookChanged = true;
                }

                int workbookViewCount = workbookViews.Elements<WorkbookView>().Count();
                int sheetCount = workbook.Sheets?.Elements<Sheet>().Count() ?? 0;
                foreach (var workbookView in workbookViews.Elements<WorkbookView>()) {
                    if (sheetCount > 0 && workbookView.ActiveTab?.Value >= sheetCount) {
                        workbookView.ActiveTab = (uint)(sheetCount - 1);
                        workbookChanged = true;
                    }
                }

                foreach (var worksheetPart in _workBookPart.WorksheetParts) {
                    bool worksheetChanged = false;
                    foreach (var sheetView in worksheetPart.Worksheet.Descendants<SheetView>()) {
                        if (sheetView.WorkbookViewId == null || sheetView.WorkbookViewId.Value >= workbookViewCount) {
                            sheetView.WorkbookViewId = 0U;
                            worksheetChanged = true;
                        }
                    }

                    if (worksheetChanged) {
                        worksheetPart.Worksheet.Save();
                    }
                }
            } else if (workbookViews != null && !workbookViews.Elements<WorkbookView>().Any()) {
                workbook.RemoveChild(workbookViews);
                workbookChanged = true;
            }

            if (save && workbookChanged) {
                workbook.Save();
            }
        }
    }
}
