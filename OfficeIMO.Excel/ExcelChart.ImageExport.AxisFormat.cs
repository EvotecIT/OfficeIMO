using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        private const int ImageExportMaxAxisNumberFormatLength = 1024;

        private string? GetImageExportHorizontalAxisNumberFormat(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            string? valueAxisNumberFormat = GetImageExportAxisNumberFormat(valueAxis);
            if (HasHorizontalBarChart(plotArea)) {
                return valueAxisNumberFormat;
            }

            return categoryAxis.Is(C, "valAx") ? GetImageExportAxisNumberFormat(categoryAxis) : null;
        }

        private string? GetImageExportVerticalAxisNumberFormat(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisNumberFormat(valueAxis);
            }

            return categoryAxis.Is(C, "valAx") ? GetImageExportAxisNumberFormat(categoryAxis) : null;
        }

        private string? GetImageExportCategoryAxisNumberFormat(OfficeXmlNode categoryAxis) {
            if (!categoryAxis.Is(C, "catAx") && !categoryAxis.Is(C, "dateAx")) {
                return null;
            }

            string? formatCode = GetImageExportAxisNumberFormat(categoryAxis);
            return formatCode != null && IsSimpleSupportedImageExportAxisNumberFormat(formatCode) ? formatCode : null;
        }

        private static (double? Divisor, string? Label) GetImageExportHorizontalAxisDisplayUnit(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            (double? Divisor, string? Label) valueAxisDisplayUnit = GetImageExportAxisDisplayUnit(valueAxis);
            if (HasHorizontalBarChart(plotArea)) {
                return valueAxisDisplayUnit;
            }

            return categoryAxis.Is(C, "valAx") ? GetImageExportAxisDisplayUnit(categoryAxis) : default;
        }

        private static (double? Divisor, string? Label) GetImageExportVerticalAxisDisplayUnit(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisDisplayUnit(valueAxis);
            }

            return categoryAxis.Is(C, "valAx") ? GetImageExportAxisDisplayUnit(categoryAxis) : default;
        }

        private static (double? Minimum, double? Maximum, double? MajorUnit, double? MinorUnit) GetImageExportHorizontalAxisScale(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            if (HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisScale(valueAxis);
            }

            return categoryAxis.Is(C, "valAx") ? GetImageExportAxisScale(categoryAxis) : default;
        }

        private static (double? Minimum, double? Maximum, double? MajorUnit, double? MinorUnit) GetImageExportVerticalAxisScale(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisScale(valueAxis);
            }

            return categoryAxis.Is(C, "valAx") ? GetImageExportAxisScale(categoryAxis) : default;
        }

        private static OfficeChartAxisTickMark GetImageExportHorizontalAxisMajorTickMark(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            if (HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickMark(valueAxis, major: true);
            }

            return GetImageExportAxisTickMark(categoryAxis, major: true);
        }

        private static OfficeChartAxisTickMark GetImageExportVerticalAxisMajorTickMark(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickMark(valueAxis, major: true);
            }

            return GetImageExportAxisTickMark(categoryAxis, major: true);
        }

        private static OfficeChartAxisTickMark GetImageExportHorizontalAxisMinorTickMark(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            if (HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickMark(valueAxis, major: false);
            }

            return GetImageExportAxisTickMark(categoryAxis, major: false);
        }

        private static OfficeChartAxisTickMark GetImageExportVerticalAxisMinorTickMark(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickMark(valueAxis, major: false);
            }

            return GetImageExportAxisTickMark(categoryAxis, major: false);
        }

        private static OfficeChartAxisTickLabelPosition GetImageExportHorizontalAxisTickLabelPosition(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            if (HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickLabelPosition(valueAxis);
            }

            return GetImageExportAxisTickLabelPosition(categoryAxis);
        }

        private static OfficeChartAxisTickLabelPosition GetImageExportVerticalAxisTickLabelPosition(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisTickLabelPosition(valueAxis);
            }

            return GetImageExportAxisTickLabelPosition(categoryAxis);
        }

        private static bool GetImageExportReverseCategoryAxis(OfficeXmlNode categoryAxis) =>
            (categoryAxis.Is(C, "catAx") || categoryAxis.Is(C, "dateAx")) &&
            categoryAxis.Child(C, "scaling").Child(C, "orientation").Val() == "maxMin";

        private static bool HasImageExportCategoryAxisOrientation(OfficeXmlNode categoryAxis) =>
            (categoryAxis.Is(C, "catAx") || categoryAxis.Is(C, "dateAx")) &&
            categoryAxis.Child(C, "scaling").Child(C, "orientation").Val() != null;

        private string? GetImageExportAxisNumberFormat(OfficeXmlNode axis) {
            OfficeXmlNode nativeFormat = axis.Child(C, "numFmt");
            string? formatCode = nativeFormat.Attribute("formatCode");
            string? value = nativeFormat.BooleanAttribute("sourceLinked") == true ? ResolveSourceLinkedAxisNumberFormat(axis) ?? formatCode : formatCode;
            return string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
        }

        private static OfficeChartAxisTickMark GetImageExportAxisTickMark(OfficeXmlNode axis, bool major) {
            string? value = axis.Child(C, major ? "majorTickMark" : "minorTickMark").Val();
            return value switch {
                "in" => OfficeChartAxisTickMark.Inside,
                "out" => OfficeChartAxisTickMark.Outside,
                "cross" => OfficeChartAxisTickMark.Cross,
                _ => OfficeChartAxisTickMark.None
            };
        }

        private static OfficeChartAxisTickLabelPosition GetImageExportAxisTickLabelPosition(OfficeXmlNode axis) {
            string? value = axis.Child(C, "tickLblPos").Val();
            return value switch {
                "none" => OfficeChartAxisTickLabelPosition.None,
                "low" => OfficeChartAxisTickLabelPosition.Low,
                "high" => OfficeChartAxisTickLabelPosition.High,
                _ => OfficeChartAxisTickLabelPosition.NextTo
            };
        }

        private static (double? Divisor, string? Label) GetImageExportAxisDisplayUnit(OfficeXmlNode axis) =>
            ImageExportReadDisplayUnits(axis);

        private static (double? Minimum, double? Maximum, double? MajorUnit, double? MinorUnit) GetImageExportAxisScale(OfficeXmlNode axis) {
            OfficeXmlNode scaling = axis.Child(C, "scaling");
            double? minimum = TryGetFiniteAxisValue(scaling.Child(C, "min").DoubleVal(), out double minValue) ? minValue : null;
            double? maximum = TryGetFiniteAxisValue(scaling.Child(C, "max").DoubleVal(), out double maxValue) ? maxValue : null;
            double? majorUnit = TryGetPositiveAxisValue(axis.Child(C, "majorUnit").DoubleVal(), out double majorValue) ? majorValue : null;
            double? minorUnit = TryGetPositiveAxisValue(axis.Child(C, "minorUnit").DoubleVal(), out double minorValue) ? minorValue : null;
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

        private bool HasUnsupportedImageExportAxisNumberFormat(OfficeXmlNode chartSpace) {
            foreach (OfficeXmlNode plotArea in chartSpace.Descendants(C, "plotArea")) {
                foreach (OfficeXmlNode axis in plotArea.Children()) {
                    if (!axis.Is(C, "valAx") && !axis.Is(C, "catAx") && !axis.Is(C, "dateAx")) continue;
                    if (axis.Child(C, "numFmt").BooleanAttribute("sourceLinked") == true &&
                        HasSourceLinkedNumericAxisReference(plotArea, axis) && ResolveSourceLinkedAxisNumberFormat(axis) == null) return true;
                }
                OfficeXmlNode categoryAxis = ResolveImageExportCategoryAxis(plotArea);
                OfficeXmlNode valueAxis = ResolveImageExportValueAxis(plotArea);

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

        private bool HasUnsupportedImageExportCategoryAxisNumberFormat(OfficeXmlNode chartSpace) {
            foreach (OfficeXmlNode plotArea in chartSpace.Descendants(C, "plotArea")) {
                foreach (OfficeXmlNode axis in plotArea.Children(C, "catAx")) {
                    if (HasUnsupportedImageExportCategoryAxisNumberFormatCore(axis)) {
                        return true;
                    }
                }

                foreach (OfficeXmlNode axis in plotArea.Children(C, "dateAx")) {
                    if (HasUnsupportedImageExportCategoryAxisNumberFormatCore(axis)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private bool HasUnsupportedImageExportCategoryAxisNumberFormatCore(OfficeXmlNode axis) {
            string? formatCode = GetImageExportAxisNumberFormat(axis);
            return formatCode != null && !IsSimpleSupportedImageExportAxisNumberFormat(formatCode);
        }

        private static bool IsSimpleSupportedImageExportAxisNumberFormat(string formatCode) {
            if (string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            return formatCode.Length <= ImageExportMaxAxisNumberFormatLength &&
                OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSecondaryAxis.IsDrawableNumberFormat(formatCode);
        }
    }
}
