using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal void CleanupWorkbookViewArtifacts(bool save = true) {
            var workbook = WorkbookRoot;
            bool hasSheetViews = WorkbookPartRoot.WorksheetParts.Any(worksheetPart =>
                (worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing."))
                    .GetFirstChild<SheetViews>()?.Elements<SheetView>().Any() == true);

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

                foreach (var worksheetPart in WorkbookPartRoot.WorksheetParts) {
                    var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
                    bool worksheetChanged = false;
                    foreach (var sheetView in worksheet.Descendants<SheetView>()) {
                        if (sheetView.WorkbookViewId == null || sheetView.WorkbookViewId.Value >= workbookViewCount) {
                            sheetView.WorkbookViewId = 0U;
                            worksheetChanged = true;
                        }
                    }

                    if (worksheetChanged) {
                        worksheet.Save();
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
