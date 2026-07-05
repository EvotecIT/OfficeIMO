using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace OfficeIMO.Excel.Utilities {
    internal readonly struct ExcelGradientFillStopInfo {
        internal ExcelGradientFillStopInfo(double offset, string colorArgb) {
            Offset = offset;
            ColorArgb = colorArgb;
        }

        internal double Offset { get; }

        internal string ColorArgb { get; }
    }

    internal readonly struct ExcelGradientFillInfo {
        internal ExcelGradientFillInfo(IReadOnlyList<ExcelGradientFillStopInfo> stops, double degree) {
            Stops = stops;
            Degree = degree;
        }

        internal IReadOnlyList<ExcelGradientFillStopInfo> Stops { get; }

        internal string StartColorArgb => Stops[0].ColorArgb;

        internal string EndColorArgb => Stops[Stops.Count - 1].ColorArgb;

        internal double Degree { get; }
    }

    internal static class ExcelGradientFillResolver {
        private const double EndpointTolerance = 0.000001D;

        internal static bool TryResolveSimpleLinearGradient(Fill? fill, WorkbookPart? workbookPart, out ExcelGradientFillInfo gradient) {
            gradient = default;
            GradientFill? gradientFill = fill?.GradientFill;
            if (gradientFill == null || IsPathGradient(gradientFill)) {
                return false;
            }

            GradientStop[] stops = gradientFill.Elements<GradientStop>()
                .OrderBy(stop => stop.Position?.Value ?? 0D)
                .ToArray();
            if (stops.Length < 2) {
                return false;
            }

            if (!IsEndpoint(stops[0], 0D) || !IsEndpoint(stops[stops.Length - 1], 1D)) {
                return false;
            }

            var resolvedStops = new List<ExcelGradientFillStopInfo>(stops.Length);
            double previousOffset = -1D;
            for (int i = 0; i < stops.Length; i++) {
                double offset = NormalizeEndpointOffset(stops[i].Position?.Value ?? 0D, i, stops.Length);
                if (offset <= previousOffset) {
                    return false;
                }

                string? color = ResolveStopColor(stops[i], workbookPart);
                if (color == null) {
                    return false;
                }

                resolvedStops.Add(new ExcelGradientFillStopInfo(offset, color));
                previousOffset = offset;
            }

            gradient = new ExcelGradientFillInfo(resolvedStops, gradientFill.Degree?.Value ?? 0D);
            return true;
        }

        private static bool IsEndpoint(GradientStop stop, double expected) =>
            Math.Abs((stop.Position?.Value ?? 0D) - expected) <= EndpointTolerance;

        private static double NormalizeEndpointOffset(double offset, int index, int stopCount) {
            if (index == 0 && Math.Abs(offset) <= EndpointTolerance) {
                return 0D;
            }

            if (index == stopCount - 1 && Math.Abs(offset - 1D) <= EndpointTolerance) {
                return 1D;
            }

            return offset;
        }

        private static bool IsPathGradient(GradientFill gradientFill) {
            string? type = gradientFill.Type?.Value.ToString();
            return string.Equals(type, "path", StringComparison.OrdinalIgnoreCase);
        }

        private static string? ResolveStopColor(GradientStop stop, WorkbookPart? workbookPart) =>
            ExcelThemeColorResolver.Resolve(stop.GetFirstChild<Color>(), workbookPart);
    }
}
