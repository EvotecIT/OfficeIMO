using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Applies worksheet view options such as gridlines, direction, zoom, and view mode.
        /// </summary>
        /// <param name="options">Options to apply. Null properties leave the existing worksheet setting unchanged.</param>
        internal void ApplyViewOptions(ExcelWorksheetViewOptions options) {
            ArgumentNullException.ThrowIfNull(options);

            ValidateZoom(options.ZoomScale, nameof(options.ZoomScale));
            ValidateZoom(options.ZoomScaleNormal, nameof(options.ZoomScaleNormal));

            WriteWorksheetView(worksheet => {
                OfficeXmlNode sheetView = GetOrCreateSheetView(worksheet);

                if (options.ShowGridlines.HasValue) {
                    sheetView.SetBooleanAttribute("showGridLines", options.ShowGridlines.Value);
                }

                if (options.RightToLeft.HasValue) {
                    sheetView.SetBooleanAttribute("rightToLeft", options.RightToLeft.Value);
                }

                if (options.ZoomScale.HasValue) {
                    sheetView.SetUInt32Attribute("zoomScale", options.ZoomScale.Value);
                }

                if (options.ZoomScaleNormal.HasValue) {
                    sheetView.SetUInt32Attribute("zoomScaleNormal", options.ZoomScaleNormal.Value);
                }

                if (options.View.HasValue) {
                    sheetView.SetAttribute("view", FormatViewKind(options.View.Value));
                }

                return true;
            });
        }

        private static string FormatViewKind(ExcelWorksheetViewKind view) => view switch {
            ExcelWorksheetViewKind.Normal => "normal",
            ExcelWorksheetViewKind.PageBreakPreview => "pageBreakPreview",
            ExcelWorksheetViewKind.PageLayout => "pageLayout",
            _ => throw new ArgumentOutOfRangeException(nameof(view), view, "Unsupported worksheet view kind.")
        };

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
