using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private OfficeXmlNode ViewWorksheetRoot {
            get {
                CaptureOriginalDynamicSpillFingerprintIfSafe();
                return ExcelPartRef.WorksheetForMetadata(this).Root;
            }
        }

        /// <summary>
        /// Reads worksheet view metadata such as frozen panes, gridline visibility, zoom, and direction.
        /// </summary>
        internal ExcelWorksheetViewInfo ReadViewInfo() {
            if (_excelDocument.ReadDirectWorksheetView(this) is { } pending) return pending;
            if (DeferredEntry is { } entry && _excelDocument.ReadDeferredWorksheetView(entry) is { } info) return info;
            return ReadViewInfo(ViewWorksheetRoot);
        }

        internal static ExcelWorksheetViewInfo ReadViewInfo(OfficeXmlNode worksheet) =>
            ReadPrimaryView(worksheet.Child(OfficeXmlNamespaces.SpreadsheetML, "sheetViews"));

        private static ExcelWorksheetViewInfo ReadPrimaryView(OfficeXmlNode views) {
            OfficeXmlNode sheetView = views.Child(OfficeXmlNamespaces.SpreadsheetML, "sheetView");
            OfficeXmlNode pane = sheetView.Child(OfficeXmlNamespaces.SpreadsheetML, "pane");
            string? state = pane.Attribute("state");
            bool frozen = state is "frozen" or "frozenSplit";
            double? horizontal = pane.DoubleAttribute("xSplit");
            double? vertical = pane.DoubleAttribute("ySplit");

            return new ExcelWorksheetViewInfo {
                HasPane = pane.Exists,
                PaneState = state,
                HorizontalSplit = horizontal,
                VerticalSplit = vertical,
                FrozenRowCount = frozen ? ConvertFrozenSplitToInt(vertical) : 0,
                FrozenColumnCount = frozen ? ConvertFrozenSplitToInt(horizontal) : 0,
                TopLeftCell = pane.Attribute("topLeftCell") ?? sheetView.Attribute("topLeftCell"),
                ActivePane = pane.Attribute("activePane"),
                ShowGridlines = sheetView.BooleanAttribute("showGridLines") ?? true,
                ShowRowColumnHeadings = sheetView.BooleanAttribute("showRowColHeaders") ?? true,
                ShowZeros = sheetView.BooleanAttribute("showZeros") ?? true,
                RightToLeft = sheetView.BooleanAttribute("rightToLeft") ?? false,
                View = sheetView.Attribute("view"),
                ZoomScale = sheetView.UInt32Attribute("zoomScale"),
                ZoomScaleNormal = sheetView.UInt32Attribute("zoomScaleNormal")
            };
        }

        private static int ConvertFrozenSplitToInt(double? split) =>
            checked((int)Math.Round(split ?? 0D, MidpointRounding.AwayFromZero));
    }
}
