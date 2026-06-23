using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Sets the worksheet that spreadsheet applications should show when the workbook opens.
        /// </summary>
        /// <param name="sheetName">Worksheet name.</param>
        public void SetActiveWorksheet(string sheetName) {
            SetActiveWorksheet(GetSheet(sheetName));
        }

        /// <summary>
        /// Sets the worksheet that spreadsheet applications should show when the workbook opens.
        /// </summary>
        /// <param name="sheetIndex">Zero-based worksheet index.</param>
        public void SetActiveWorksheet(int sheetIndex) {
            var sheets = Sheets;
            if (sheetIndex < 0 || sheetIndex >= sheets.Count) {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), $"Index {sheetIndex.ToString(CultureInfo.InvariantCulture)} is out of range (0..{(sheets.Count - 1).ToString(CultureInfo.InvariantCulture)}).");
            }

            SetActiveWorksheet(sheets[sheetIndex]);
        }

        /// <summary>
        /// Sets the worksheet that spreadsheet applications should show when the workbook opens.
        /// </summary>
        /// <param name="sheet">Worksheet to activate.</param>
        public void SetActiveWorksheet(ExcelSheet sheet) {
            if (sheet == null) {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (!ReferenceEquals(sheet.Document, this)) {
                throw new ArgumentException("Worksheet must belong to this workbook.", nameof(sheet));
            }

            MaterializeDeferredDataSetImport();
            Locking.ExecuteWrite(EnsureLock(), () => {
                Sheets workbookSheets = WorkbookRoot.Sheets ?? throw new InvalidOperationException("Workbook sheets collection is missing.");
                List<Sheet> orderedSheets = workbookSheets.Elements<Sheet>().ToList();
                int activeIndex = orderedSheets.FindIndex(candidate => ReferenceEquals(candidate, sheet.SheetElement)
                    || string.Equals(candidate.Name?.Value, sheet.Name, StringComparison.Ordinal));
                if (activeIndex < 0) {
                    throw new ArgumentException("Worksheet not found in workbook.", nameof(sheet));
                }

                Sheet activeSheet = orderedSheets[activeIndex];
                if (activeSheet.State?.Value == SheetStateValues.Hidden || activeSheet.State?.Value == SheetStateValues.VeryHidden) {
                    throw new InvalidOperationException("A hidden worksheet cannot be the active worksheet.");
                }

                WorkbookView workbookView = GetOrCreatePrimaryWorkbookViewForActiveWorksheet();
                workbookView.ActiveTab = (uint)activeIndex;
                workbookView.FirstSheet = (uint)activeIndex;

                for (int index = 0; index < orderedSheets.Count; index++) {
                    Sheet currentSheet = orderedSheets[index];
                    if (currentSheet.Id?.Value == null || WorkbookPartRoot.GetPartById(currentSheet.Id.Value) is not WorksheetPart worksheetPart) {
                        continue;
                    }

                    Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
                    SheetView sheetView = GetOrCreatePrimarySheetViewForActiveWorksheet(worksheet);
                    sheetView.WorkbookViewId ??= 0U;
                    sheetView.TabSelected = index == activeIndex;
                    worksheet.Save();
                }

                WorkbookRoot.Save();
                MarkPackageDirty();
            });
        }

        private WorkbookView GetOrCreatePrimaryWorkbookViewForActiveWorksheet() {
            BookViews? workbookViews = WorkbookRoot.GetFirstChild<BookViews>();
            if (workbookViews == null) {
                workbookViews = new BookViews();
                var sheets = WorkbookRoot.GetFirstChild<Sheets>();
                if (sheets != null) {
                    WorkbookRoot.InsertBefore(workbookViews, sheets);
                } else {
                    WorkbookRoot.Append(workbookViews);
                }
            }

            WorkbookView? workbookView = workbookViews.GetFirstChild<WorkbookView>();
            if (workbookView == null) {
                workbookView = new WorkbookView();
                workbookViews.Append(workbookView);
            }

            return workbookView;
        }

        private static SheetView GetOrCreatePrimarySheetViewForActiveWorksheet(Worksheet worksheet) {
            SheetViews? sheetViews = worksheet.GetFirstChild<SheetViews>();
            if (sheetViews == null) {
                sheetViews = new SheetViews();
                var sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData != null) {
                    worksheet.InsertBefore(sheetViews, sheetData);
                } else {
                    worksheet.Append(sheetViews);
                }
            }

            SheetView? sheetView = sheetViews.GetFirstChild<SheetView>();
            if (sheetView == null) {
                sheetView = new SheetView { WorkbookViewId = 0U };
                sheetViews.Append(sheetView);
            }

            return sheetView;
        }
    }
}
