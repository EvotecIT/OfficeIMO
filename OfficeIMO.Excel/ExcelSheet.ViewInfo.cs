using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Reads worksheet view metadata such as frozen panes, gridline visibility, zoom, and direction.
        /// </summary>
        public ExcelWorksheetViewInfo GetViewInfo() {
            SheetView? sheetView = WorksheetRoot
                .GetFirstChild<SheetViews>()?
                .Elements<SheetView>()
                .FirstOrDefault();
            Pane? pane = sheetView?.GetFirstChild<Pane>();
            bool frozen = pane != null && IsFrozenPaneState(pane.State?.Value);

            return new ExcelWorksheetViewInfo {
                HasPane = pane != null,
                PaneState = pane?.State?.InnerText,
                HorizontalSplit = pane?.HorizontalSplit?.Value,
                VerticalSplit = pane?.VerticalSplit?.Value,
                FrozenRowCount = frozen ? ConvertFrozenSplitToInt(pane?.VerticalSplit) : 0,
                FrozenColumnCount = frozen ? ConvertFrozenSplitToInt(pane?.HorizontalSplit) : 0,
                TopLeftCell = pane?.TopLeftCell?.Value ?? sheetView?.TopLeftCell?.Value,
                ActivePane = pane?.ActivePane?.InnerText,
                ShowGridlines = sheetView?.ShowGridLines?.Value ?? true,
                RightToLeft = sheetView?.RightToLeft?.Value ?? false,
                View = sheetView?.View?.InnerText,
                ZoomScale = sheetView?.ZoomScale?.Value,
                ZoomScaleNormal = sheetView?.ZoomScaleNormal?.Value
            };
        }

        private static int ConvertFrozenSplitToInt(DoubleValue? split) {
            if (split == null) {
                return 0;
            }

            return checked((int)Math.Round(split.Value, MidpointRounding.AwayFromZero));
        }
    }
}
