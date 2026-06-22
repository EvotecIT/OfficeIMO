using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Applies worksheet view options such as gridlines, direction, zoom, and view mode.
        /// </summary>
        /// <param name="options">Options to apply. Null properties leave the existing worksheet setting unchanged.</param>
        public void SetViewOptions(ExcelWorksheetViewOptions options) {
            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            ValidateZoom(options.ZoomScale, nameof(options.ZoomScale));
            ValidateZoom(options.ZoomScaleNormal, nameof(options.ZoomScaleNormal));

            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetView sheetView = GetOrCreatePrimarySheetViewForViewOptions();

                if (options.ShowGridlines.HasValue) {
                    sheetView.ShowGridLines = options.ShowGridlines.Value;
                }

                if (options.RightToLeft.HasValue) {
                    sheetView.RightToLeft = options.RightToLeft.Value;
                }

                if (options.ZoomScale.HasValue) {
                    sheetView.ZoomScale = options.ZoomScale.Value;
                }

                if (options.ZoomScaleNormal.HasValue) {
                    sheetView.ZoomScaleNormal = options.ZoomScaleNormal.Value;
                }

                if (options.View.HasValue) {
                    sheetView.View = ToOpenXmlSheetView(options.View.Value);
                }

                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Applies worksheet view options such as gridlines, direction, zoom, and view mode.
        /// </summary>
        public void SetViewOptions(bool? showGridlines = null, bool? rightToLeft = null, uint? zoomScale = null, uint? zoomScaleNormal = null, ExcelWorksheetViewKind? view = null) {
            SetViewOptions(new ExcelWorksheetViewOptions {
                ShowGridlines = showGridlines,
                RightToLeft = rightToLeft,
                ZoomScale = zoomScale,
                ZoomScaleNormal = zoomScaleNormal,
                View = view,
            });
        }

        private SheetView GetOrCreatePrimarySheetViewForViewOptions() {
            var worksheet = WorksheetRoot;
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
            } else if (sheetView.WorkbookViewId == null) {
                sheetView.WorkbookViewId = 0U;
            }

            return sheetView;
        }

        private static SheetViewValues ToOpenXmlSheetView(ExcelWorksheetViewKind view) {
            return view switch {
                ExcelWorksheetViewKind.Normal => SheetViewValues.Normal,
                ExcelWorksheetViewKind.PageBreakPreview => SheetViewValues.PageBreakPreview,
                ExcelWorksheetViewKind.PageLayout => SheetViewValues.PageLayout,
                _ => throw new ArgumentOutOfRangeException(nameof(view), view, "Unsupported worksheet view kind."),
            };
        }

        private static void ValidateZoom(uint? zoom, string parameterName) {
            if (!zoom.HasValue) {
                return;
            }

            if (zoom.Value < 10U || zoom.Value > 400U) {
                throw new ArgumentOutOfRangeException(parameterName, zoom.Value, "Worksheet zoom must be between 10 and 400.");
            }
        }
    }
}
