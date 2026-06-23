using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Utilities {
    internal readonly struct ExcelGradientFillInfo {
        internal ExcelGradientFillInfo(string startColorArgb, string endColorArgb, double degree) {
            StartColorArgb = startColorArgb;
            EndColorArgb = endColorArgb;
            Degree = degree;
        }

        internal string StartColorArgb { get; }

        internal string EndColorArgb { get; }

        internal double Degree { get; }
    }

    internal static class ExcelGradientFillResolver {
        internal static bool TryResolveSimpleLinearGradient(Fill? fill, WorkbookPart? workbookPart, out ExcelGradientFillInfo gradient) {
            gradient = default;
            GradientFill? gradientFill = fill?.GradientFill;
            if (gradientFill == null || IsPathGradient(gradientFill)) {
                return false;
            }

            GradientStop[] stops = gradientFill.Elements<GradientStop>()
                .OrderBy(stop => stop.Position?.Value ?? 0D)
                .ToArray();
            if (stops.Length != 2) {
                return false;
            }

            string? start = ResolveStopColor(stops[0], workbookPart);
            string? end = ResolveStopColor(stops[1], workbookPart);
            if (start == null || end == null) {
                return false;
            }

            gradient = new ExcelGradientFillInfo(start, end, gradientFill.Degree?.Value ?? 0D);
            return true;
        }

        private static bool IsPathGradient(GradientFill gradientFill) {
            string? type = gradientFill.Type?.Value.ToString();
            return string.Equals(type, "path", StringComparison.OrdinalIgnoreCase);
        }

        private static string? ResolveStopColor(GradientStop stop, WorkbookPart? workbookPart) =>
            ExcelThemeColorResolver.Resolve(stop.GetFirstChild<Color>(), workbookPart);
    }
}
