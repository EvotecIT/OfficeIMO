using System.Globalization;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        private const int ImageExportMaxAxisNumberFormatLength = 1024;

        private string? GetImageExportHorizontalAxisNumberFormat(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            string? valueAxisNumberFormat = GetImageExportAxisNumberFormat(valueAxis);
            if (HasHorizontalBarChart(plotArea)) {
                return valueAxisNumberFormat;
            }

            return categoryAxis is C.ValueAxis ? GetImageExportAxisNumberFormat(categoryAxis) : null;
        }

        private string? GetImageExportVerticalAxisNumberFormat(C.PlotArea plotArea, OpenXmlCompositeElement? categoryAxis, OpenXmlCompositeElement? valueAxis) {
            if (!HasHorizontalBarChart(plotArea)) {
                return GetImageExportAxisNumberFormat(valueAxis);
            }

            return categoryAxis is C.ValueAxis ? GetImageExportAxisNumberFormat(categoryAxis) : null;
        }

        private string? GetImageExportCategoryAxisNumberFormat(OpenXmlCompositeElement? categoryAxis) {
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

        private static bool GetImageExportReverseCategoryAxis(OpenXmlCompositeElement? categoryAxis) =>
            (categoryAxis is C.CategoryAxis || categoryAxis is C.DateAxis) &&
            categoryAxis.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin;

        private static bool HasImageExportCategoryAxisOrientation(OpenXmlCompositeElement? categoryAxis) =>
            (categoryAxis is C.CategoryAxis || categoryAxis is C.DateAxis) &&
            categoryAxis.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>()?.Val != null;

        private string? GetImageExportAxisNumberFormat(OpenXmlCompositeElement? axis) {
            C.NumberingFormat? nativeFormat = axis?.GetFirstChild<C.NumberingFormat>();
            string? value = nativeFormat?.SourceLinked?.Value == true ? ResolveSourceLinkedAxisNumberFormat(axis!) ?? nativeFormat.FormatCode?.Value : nativeFormat?.FormatCode?.Value;
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

        private static (double? Divisor, string? Label) GetImageExportAxisDisplayUnit(OpenXmlCompositeElement? axis) =>
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartPlotSemantics.ReadDisplayUnits(axis);

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

        private bool HasUnsupportedImageExportAxisNumberFormat(C.ChartSpace chartSpace) {
            foreach (C.PlotArea plotArea in chartSpace.Descendants<C.PlotArea>()) {
                foreach (var axis in plotArea.ChildElements.OfType<OpenXmlCompositeElement>().Where(axis => axis is C.ValueAxis or C.CategoryAxis or C.DateAxis)) {
                    if (axis.GetFirstChild<C.NumberingFormat>()?.SourceLinked?.Value == true &&
                        HasSourceLinkedNumericAxisReference(plotArea, axis) && ResolveSourceLinkedAxisNumberFormat(axis) == null) return true;
                }
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

        private bool HasUnsupportedImageExportCategoryAxisNumberFormat(C.ChartSpace chartSpace) {
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

        private bool HasUnsupportedImageExportCategoryAxisNumberFormat(OpenXmlCompositeElement axis) {
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
