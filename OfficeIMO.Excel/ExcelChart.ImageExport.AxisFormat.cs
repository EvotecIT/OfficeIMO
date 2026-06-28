using System;
using System.Globalization;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        private const int ImageExportMaxAxisNumberFormatLength = 1024;

        private static string? GetImageExportHorizontalAxisNumberFormat(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            string? valueAxisNumberFormat = GetImageExportAxisNumberFormat(valueAxis);
            if (HasHorizontalBarChart(plotArea)) {
                return valueAxisNumberFormat;
            }

            return categoryAxis is C.ValueAxis ? GetImageExportAxisNumberFormat(categoryAxis) : null;
        }

        private static string? GetImageExportVerticalAxisNumberFormat(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisNumberFormat(valueAxis);
            }

            return categoryAxis is C.ValueAxis ? GetImageExportAxisNumberFormat(categoryAxis) : null;
        }

        private static string? GetImageExportCategoryAxisNumberFormat(OpenXmlCompositeElement? categoryAxis) {
            if (categoryAxis is not C.CategoryAxis && categoryAxis is not C.DateAxis) {
                return null;
            }

            string? formatCode = GetImageExportAxisNumberFormat(categoryAxis);
            return formatCode != null && IsSimpleSupportedImageExportAxisNumberFormat(formatCode) ? formatCode : null;
        }

        private static (double? Divisor, string? Label) GetImageExportHorizontalAxisDisplayUnit(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            (double? Divisor, string? Label) valueAxisDisplayUnit = GetImageExportAxisDisplayUnit(valueAxis);
            if (HasHorizontalBarChart(plotArea)) {
                return valueAxisDisplayUnit;
            }

            return categoryAxis is C.ValueAxis ? GetImageExportAxisDisplayUnit(categoryAxis) : default;
        }

        private static (double? Divisor, string? Label) GetImageExportVerticalAxisDisplayUnit(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisDisplayUnit(valueAxis);
            }

            return categoryAxis is C.ValueAxis ? GetImageExportAxisDisplayUnit(categoryAxis) : default;
        }

        private static (double? Minimum, double? Maximum, double? MajorUnit, double? MinorUnit) GetImageExportHorizontalAxisScale(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            if (HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisScale(valueAxis);
            }

            return categoryAxis is C.ValueAxis ? GetImageExportAxisScale(categoryAxis) : default;
        }

        private static (double? Minimum, double? Maximum, double? MajorUnit, double? MinorUnit) GetImageExportVerticalAxisScale(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisScale(valueAxis);
            }

            return categoryAxis is C.ValueAxis ? GetImageExportAxisScale(categoryAxis) : default;
        }

        private static OfficeChartAxisTickMark GetImageExportHorizontalAxisMajorTickMark(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            if (HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickMark(valueAxis, major: true);
            }

            return GetImageExportAxisTickMark(categoryAxis, major: true);
        }

        private static OfficeChartAxisTickMark GetImageExportVerticalAxisMajorTickMark(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickMark(valueAxis, major: true);
            }

            return GetImageExportAxisTickMark(categoryAxis, major: true);
        }

        private static OfficeChartAxisTickMark GetImageExportHorizontalAxisMinorTickMark(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            if (HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickMark(valueAxis, major: false);
            }

            return GetImageExportAxisTickMark(categoryAxis, major: false);
        }

        private static OfficeChartAxisTickMark GetImageExportVerticalAxisMinorTickMark(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickMark(valueAxis, major: false);
            }

            return GetImageExportAxisTickMark(categoryAxis, major: false);
        }

        private static OfficeChartAxisTickLabelPosition GetImageExportHorizontalAxisTickLabelPosition(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            if (HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickLabelPosition(valueAxis);
            }

            return GetImageExportAxisTickLabelPosition(categoryAxis);
        }

        private static OfficeChartAxisTickLabelPosition GetImageExportVerticalAxisTickLabelPosition(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickLabelPosition(valueAxis);
            }

            return GetImageExportAxisTickLabelPosition(categoryAxis);
        }

        private static OfficeChartAxisCrossingPosition GetImageExportVerticalAxisCrossingPosition(C.PlotArea plotArea, OpenXmlCompositeElement? valueAxis) {
            if (HasHorizontalBarChart(plotArea)) {
                return OfficeChartAxisCrossingPosition.AutoZero;
            }

            return GetImageExportAxisCrossingPosition(valueAxis);
        }

        private static OfficeChartAxisCrossingPosition GetImageExportHorizontalAxisCrossingPosition(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis) {
            if (HasHorizontalBarChart(plotArea)) {
                return OfficeChartAxisCrossingPosition.AutoZero;
            }

            return GetImageExportAxisCrossingPosition(categoryAxis);
        }

        private static bool GetImageExportReverseCategoryAxis(OpenXmlCompositeElement? categoryAxis) =>
            (categoryAxis is C.CategoryAxis || categoryAxis is C.DateAxis) &&
            categoryAxis.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin;

        private static bool HasImageExportCategoryAxisOrientation(OpenXmlCompositeElement? categoryAxis) =>
            (categoryAxis is C.CategoryAxis || categoryAxis is C.DateAxis) &&
            categoryAxis.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>()?.Val != null;

        private static string? GetImageExportAxisNumberFormat(OpenXmlCompositeElement? axis) {
            string? value = axis?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value;
            return string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
        }

        private static OfficeChartAxisTickMark GetImageExportAxisTickMark(OpenXmlCompositeElement? axis, bool major) {
            C.TickMarkValues? value = major
                ? axis?.GetFirstChild<C.MajorTickMark>()?.Val?.Value
                : axis?.GetFirstChild<C.MinorTickMark>()?.Val?.Value;
            if (value == null || value.Value == C.TickMarkValues.None) {
                return OfficeChartAxisTickMark.None;
            }

            if (value.Value == C.TickMarkValues.Inside) {
                return OfficeChartAxisTickMark.Inside;
            }

            if (value.Value == C.TickMarkValues.Outside) {
                return OfficeChartAxisTickMark.Outside;
            }

            return value.Value == C.TickMarkValues.Cross
                ? OfficeChartAxisTickMark.Cross
                : OfficeChartAxisTickMark.None;
        }

        private static OfficeChartAxisTickLabelPosition GetImageExportAxisTickLabelPosition(OpenXmlCompositeElement? axis) {
            C.TickLabelPositionValues? value = axis?.GetFirstChild<C.TickLabelPosition>()?.Val?.Value;
            if (value == null || value.Value == C.TickLabelPositionValues.NextTo) {
                return OfficeChartAxisTickLabelPosition.NextTo;
            }

            if (value.Value == C.TickLabelPositionValues.None) {
                return OfficeChartAxisTickLabelPosition.None;
            }

            if (value.Value == C.TickLabelPositionValues.Low) {
                return OfficeChartAxisTickLabelPosition.Low;
            }

            return value.Value == C.TickLabelPositionValues.High
                ? OfficeChartAxisTickLabelPosition.High
                : OfficeChartAxisTickLabelPosition.NextTo;
        }

        private static OfficeChartAxisCrossingPosition GetImageExportAxisCrossingPosition(OpenXmlCompositeElement? axis) {
            C.CrossesValues? value = axis?.GetFirstChild<C.Crosses>()?.Val?.Value;
            return value == C.CrossesValues.Maximum
                ? OfficeChartAxisCrossingPosition.Maximum
                : OfficeChartAxisCrossingPosition.AutoZero;
        }

        private static (double? Divisor, string? Label) GetImageExportAxisDisplayUnit(OpenXmlCompositeElement? axis) {
            C.DisplayUnits? displayUnits = axis?.GetFirstChild<C.DisplayUnits>();
            if (displayUnits == null) {
                return default;
            }

            C.BuiltInUnit? builtIn = displayUnits.GetFirstChild<C.BuiltInUnit>();
            string? builtInUnit = builtIn?.Val?.InnerText ?? builtIn?.Val?.Value.ToString() ?? builtIn?.InnerText;
            double? divisor = builtInUnit == null
                ? displayUnits.GetFirstChild<C.CustomDisplayUnit>()?.Val?.Value
                : GetImageExportBuiltInDisplayUnitDivisor(builtInUnit);
            if (!divisor.HasValue || divisor.Value <= 0D || double.IsNaN(divisor.Value) || double.IsInfinity(divisor.Value)) {
                return default;
            }

            C.DisplayUnitsLabel? displayUnitsLabel = displayUnits.GetFirstChild<C.DisplayUnitsLabel>();
            string? label = displayUnitsLabel?.GetFirstChild<C.ChartText>()?.InnerText;
            if (displayUnitsLabel != null && string.IsNullOrWhiteSpace(label)) {
                label = builtInUnit == null
                    ? divisor.Value.ToString("0.##", CultureInfo.InvariantCulture)
                    : GetImageExportBuiltInDisplayUnitLabel(builtInUnit);
            }

            return (divisor.Value, string.IsNullOrWhiteSpace(label) ? null : label!.Trim());
        }

        private static (double? Minimum, double? Maximum, double? MajorUnit, double? MinorUnit) GetImageExportAxisScale(OpenXmlCompositeElement? axis) {
            C.Scaling? scaling = axis?.GetFirstChild<C.Scaling>();
            double? minimum = TryGetFiniteAxisValue(scaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value, out double minValue) ? minValue : null;
            double? maximum = TryGetFiniteAxisValue(scaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value, out double maxValue) ? maxValue : null;
            double? majorUnit = TryGetPositiveAxisValue(axis?.GetFirstChild<C.MajorUnit>()?.Val?.Value, out double majorValue) ? majorValue : null;
            double? minorUnit = TryGetPositiveAxisValue(axis?.GetFirstChild<C.MinorUnit>()?.Val?.Value, out double minorValue) ? minorValue : null;
            return (minimum, maximum, majorUnit, minorUnit);
        }

        private static bool TryGetFiniteAxisValue(double? value, out double result) {
            result = value.GetValueOrDefault();
            return value.HasValue && !double.IsNaN(result) && !double.IsInfinity(result);
        }

        private static bool TryGetPositiveAxisValue(double? value, out double result) {
            if (TryGetFiniteAxisValue(value, out result) && result > 0D) {
                return true;
            }

            result = default;
            return false;
        }

        private static double? GetImageExportBuiltInDisplayUnitDivisor(string builtInUnit) =>
            NormalizeImageExportBuiltInDisplayUnit(builtInUnit) switch {
                "hundreds" => 100D,
                "thousands" => 1_000D,
                "tenthousands" => 10_000D,
                "hundredthousands" => 100_000D,
                "millions" => 1_000_000D,
                "tenmillions" => 10_000_000D,
                "hundredmillions" => 100_000_000D,
                "billions" => 1_000_000_000D,
                "trillions" => 1_000_000_000_000D,
                _ => null
            };

        private static string GetImageExportBuiltInDisplayUnitLabel(string builtInUnit) =>
            NormalizeImageExportBuiltInDisplayUnit(builtInUnit) switch {
                "hundreds" => "Hundreds",
                "thousands" => "Thousands",
                "tenthousands" => "Ten Thousands",
                "hundredthousands" => "Hundred Thousands",
                "millions" => "Millions",
                "tenmillions" => "Ten Millions",
                "hundredmillions" => "Hundred Millions",
                "billions" => "Billions",
                "trillions" => "Trillions",
                _ => builtInUnit
            };

        private static string NormalizeImageExportBuiltInDisplayUnit(string builtInUnit) =>
            builtInUnit.Replace(" ", string.Empty)
                .Replace("-", string.Empty)
                .Replace("_", string.Empty)
                .ToLowerInvariant();

        private static bool HasUnsupportedImageExportAxisNumberFormat(C.ChartSpace chartSpace) {
            foreach (C.PlotArea plotArea in chartSpace.Descendants<C.PlotArea>()) {
                OpenXmlCompositeElement? categoryAxis = ResolveImageExportCategoryAxis(plotArea);
                OpenXmlCompositeElement? valueAxis = ResolveImageExportValueAxis(plotArea);

                string? horizontalFormatCode = GetImageExportHorizontalAxisNumberFormat(plotArea, categoryAxis, valueAxis);
                if (horizontalFormatCode != null && !IsSimpleSupportedImageExportAxisNumberFormat(horizontalFormatCode)) {
                    return true;
                }

                string? verticalFormatCode = GetImageExportVerticalAxisNumberFormat(plotArea, categoryAxis, valueAxis);
                if (verticalFormatCode != null && !IsSimpleSupportedImageExportAxisNumberFormat(verticalFormatCode)) {
                    return true;
                }
            }

            return false;
        }

        private static bool HasUnsupportedImageExportCategoryAxisNumberFormat(C.ChartSpace chartSpace) {
            foreach (C.PlotArea plotArea in chartSpace.Descendants<C.PlotArea>()) {
                foreach (C.CategoryAxis axis in plotArea.Elements<C.CategoryAxis>()) {
                    if (HasUnsupportedImageExportCategoryAxisNumberFormat(axis)) {
                        return true;
                    }
                }

                foreach (C.DateAxis axis in plotArea.Elements<C.DateAxis>()) {
                    if (HasUnsupportedImageExportCategoryAxisNumberFormat(axis)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasUnsupportedImageExportCategoryAxisNumberFormat(OpenXmlCompositeElement axis) {
            string? formatCode = GetImageExportAxisNumberFormat(axis);
            return formatCode != null && !IsSimpleSupportedImageExportAxisNumberFormat(formatCode);
        }

        private static bool IsSimpleSupportedImageExportAxisNumberFormat(string formatCode) {
            if (string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (formatCode.Length > ImageExportMaxAxisNumberFormatLength ||
                formatCode.IndexOf('[') >= 0 ||
                formatCode.IndexOf('@') >= 0) {
                return false;
            }

            for (int i = 0; i < formatCode.Length; i++) {
                if (char.IsLetter(formatCode[i])) {
                    return false;
                }
            }

            return true;
        }
    }
}
