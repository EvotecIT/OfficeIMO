using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a chart on a worksheet.
    /// </summary>
    public sealed partial class ExcelChart {
        /// <summary>
        /// Sets the category axis title.
        /// </summary>
        public ExcelChart SetCategoryAxisTitle(string title) {
            return SetCategoryAxisTitle(title, ExcelChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the category axis title for the selected axis group.
        /// </summary>
        public ExcelChart SetCategoryAxisTitle(string title, ExcelChartAxisGroup axisGroup) {
            return SetAxisTitle(title, axisGroup, AxisKind.Category);
        }

        /// <summary>
        /// Sets the value axis title.
        /// </summary>
        public ExcelChart SetValueAxisTitle(string title) {
            return SetValueAxisTitle(title, ExcelChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the value axis title for the selected axis group.
        /// </summary>
        public ExcelChart SetValueAxisTitle(string title, ExcelChartAxisGroup axisGroup) {
            return SetAxisTitle(title, axisGroup, AxisKind.Value);
        }

        /// <summary>
        /// Sets the category axis title text style.
        /// </summary>
        public ExcelChart SetCategoryAxisTitleTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisTitleTextStyle(axisGroup, AxisKind.Category, fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        /// Sets the value axis title text style.
        /// </summary>
        public ExcelChart SetValueAxisTitleTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisTitleTextStyle(axisGroup, AxisKind.Value, fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        /// Sets category axis gridlines visibility and optional styling.
        /// </summary>
        public ExcelChart SetCategoryAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisGridlines(axisGroup, AxisKind.Category, showMajor, showMinor, lineColor, lineWidthPoints);
        }

        /// <summary>
        /// Sets value axis gridlines visibility and optional styling.
        /// </summary>
        public ExcelChart SetValueAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisGridlines(axisGroup, AxisKind.Value, showMajor, showMinor, lineColor, lineWidthPoints);
        }

        /// <summary>
        /// Sets the category axis label text style.
        /// </summary>
        public ExcelChart SetCategoryAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisLabelTextStyle(axisGroup, AxisKind.Category, fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        /// Sets the value axis label text style.
        /// </summary>
        public ExcelChart SetValueAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisLabelTextStyle(axisGroup, AxisKind.Value, fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        /// Sets the category axis label rotation in degrees (-90..90).
        /// </summary>
        public ExcelChart SetCategoryAxisLabelRotation(double rotationDegrees,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisLabelRotation(axisGroup, AxisKind.Category, rotationDegrees);
        }

        /// <summary>
        /// Sets the value axis label rotation in degrees (-90..90).
        /// </summary>
        public ExcelChart SetValueAxisLabelRotation(double rotationDegrees,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisLabelRotation(axisGroup, AxisKind.Value, rotationDegrees);
        }

        /// <summary>
        /// Sets the category axis tick label position.
        /// </summary>
        public ExcelChart SetCategoryAxisTickLabelPosition(C.TickLabelPositionValues position,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisTickLabelPosition(axisGroup, AxisKind.Category, position);
        }

        /// <summary>
        /// Sets the value axis tick label position.
        /// </summary>
        public ExcelChart SetValueAxisTickLabelPosition(C.TickLabelPositionValues position,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisTickLabelPosition(axisGroup, AxisKind.Value, position);
        }

        /// <summary>
        /// Sets how the value axis crosses between categories.
        /// </summary>
        public ExcelChart SetValueAxisCrossBetween(C.CrossBetweenValues between,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            ReplaceChild(axis, new C.CrossBetween { Val = between });
            Save();
            return this;
        }

        /// <summary>
        /// Sets where the category axis crosses the value axis.
        /// </summary>
        public ExcelChart SetCategoryAxisCrossing(C.CrossesValues crosses, double? crossesAt = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.CategoryAxis? axis = ResolveCategoryAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            ApplyAxisCrossing(axis, crosses, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        /// Sets where the value axis crosses the category axis.
        /// </summary>
        public ExcelChart SetValueAxisCrossing(C.CrossesValues crosses, double? crossesAt = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            ValidateCrossesAtForAxis(axis, crossesAt);
            ApplyAxisCrossing(axis, crosses, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        /// Sets where the scatter X-axis crosses the Y-axis.
        /// </summary>
        public ExcelChart SetScatterXAxisCrossing(C.CrossesValues crosses, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveScatterXAxis(plotArea);
            if (axis == null) {
                return this;
            }

            ValidateCrossesAtForAxis(axis, crossesAt);
            ApplyAxisCrossing(axis, crosses, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        /// Sets where the scatter Y-axis crosses the X-axis.
        /// </summary>
        public ExcelChart SetScatterYAxisCrossing(C.CrossesValues crosses, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveScatterYAxis(plotArea);
            if (axis == null) {
                return this;
            }

            ValidateCrossesAtForAxis(axis, crossesAt);
            ApplyAxisCrossing(axis, crosses, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        /// Sets display units for the value axis.
        /// </summary>
        public ExcelChart SetValueAxisDisplayUnits(C.BuiltInUnitValues unit, bool showLabel = true,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            C.DisplayUnits displayUnits = axis.GetFirstChild<C.DisplayUnits>() ?? new C.DisplayUnits();
            displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
            displayUnits.RemoveAllChildren<C.BuiltInUnit>();
            displayUnits.Append(new C.BuiltInUnit { Val = unit });
            ApplyDisplayUnitsLabel(displayUnits, showLabel);
            if (displayUnits.Parent == null) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets display units for the value axis with custom label text.
        /// </summary>
        public ExcelChart SetValueAxisDisplayUnits(C.BuiltInUnitValues unit, string labelText, bool showLabel = true,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            C.DisplayUnits displayUnits = axis.GetFirstChild<C.DisplayUnits>() ?? new C.DisplayUnits();
            displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
            displayUnits.RemoveAllChildren<C.BuiltInUnit>();
            displayUnits.Append(new C.BuiltInUnit { Val = unit });
            ApplyDisplayUnitsLabel(displayUnits, showLabel, labelText);
            if (displayUnits.Parent == null) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets custom display units for the value axis.
        /// </summary>
        public ExcelChart SetValueAxisDisplayUnits(double customUnit, bool showLabel = true,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            if (customUnit <= 0 || double.IsNaN(customUnit) || double.IsInfinity(customUnit)) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            C.DisplayUnits displayUnits = axis.GetFirstChild<C.DisplayUnits>() ?? new C.DisplayUnits();
            displayUnits.RemoveAllChildren<C.BuiltInUnit>();
            displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
            displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            ApplyDisplayUnitsLabel(displayUnits, showLabel);
            if (displayUnits.Parent == null) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets custom display units for the value axis with custom label text.
        /// </summary>
        public ExcelChart SetValueAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            if (customUnit <= 0 || double.IsNaN(customUnit) || double.IsInfinity(customUnit)) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            C.DisplayUnits displayUnits = axis.GetFirstChild<C.DisplayUnits>() ?? new C.DisplayUnits();
            displayUnits.RemoveAllChildren<C.BuiltInUnit>();
            displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
            displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            ApplyDisplayUnitsLabel(displayUnits, showLabel, labelText);
            if (displayUnits.Parent == null) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Clears display units from the value axis.
        /// </summary>
        public ExcelChart ClearValueAxisDisplayUnits(ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            axis.GetFirstChild<C.DisplayUnits>()?.Remove();
            Save();
            return this;
        }

        /// <summary>
        /// Sets the category axis orientation (normal or reversed order).
        /// </summary>
        public ExcelChart SetCategoryAxisReverseOrder(bool reverseOrder = true,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.CategoryAxis? axis = ResolveCategoryAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            C.Scaling scaling = EnsureScaling(axis);
            ReplaceChild(scaling, new C.Orientation {
                Val = reverseOrder ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax
            });
            Save();
            return this;
        }

        /// <summary>
        /// Sets scatter chart X-axis scale (value axis on the bottom).
        /// </summary>
        public ExcelChart SetScatterXAxisScale(double? minimum = null, double? maximum = null,
            double? majorUnit = null, double? minorUnit = null, bool? reverseOrder = null,
            bool? logScale = null, double? logBase = null) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale, logBase);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveScatterXAxis(plotArea);
            if (axis == null) {
                return this;
            }

            ApplyAxisScale(axis, minimum, maximum, majorUnit, minorUnit, reverseOrder, logScale, logBase);
            Save();
            return this;
        }

        /// <summary>
        /// Sets scatter chart Y-axis scale (value axis on the left).
        /// </summary>
        public ExcelChart SetScatterYAxisScale(double? minimum = null, double? maximum = null,
            double? majorUnit = null, double? minorUnit = null, bool? reverseOrder = null,
            bool? logScale = null, double? logBase = null) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale, logBase);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveScatterYAxis(plotArea);
            if (axis == null) {
                return this;
            }

            ApplyAxisScale(axis, minimum, maximum, majorUnit, minorUnit, reverseOrder, logScale, logBase);
            Save();
            return this;
        }

        /// <summary>
        /// Sets the category axis number format.
        /// </summary>
        public ExcelChart SetCategoryAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetCategoryAxisNumberFormat(formatCode, sourceLinked, ExcelChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the category axis number format for the selected axis group.
        /// </summary>
        public ExcelChart SetCategoryAxisNumberFormat(string formatCode, bool sourceLinked, ExcelChartAxisGroup axisGroup) {
            return SetAxisNumberFormat(formatCode, sourceLinked, axisGroup, AxisKind.Category);
        }

        /// <summary>
        /// Sets the value axis number format.
        /// </summary>
        public ExcelChart SetValueAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetValueAxisNumberFormat(formatCode, sourceLinked, ExcelChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the value axis number format for the selected axis group.
        /// </summary>
        public ExcelChart SetValueAxisNumberFormat(string formatCode, bool sourceLinked, ExcelChartAxisGroup axisGroup) {
            return SetAxisNumberFormat(formatCode, sourceLinked, axisGroup, AxisKind.Value);
        }

        /// <summary>
        /// Sets value axis scale parameters for the selected axis group.
        /// </summary>
        public ExcelChart SetValueAxisScale(double? minimum = null, double? maximum = null, double? majorUnit = null,
            double? minorUnit = null, double? logBase = null, bool? reverseOrder = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary, bool? logScale = null) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale, logBase);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            ApplyAxisScale(axis, minimum, maximum, majorUnit, minorUnit, reverseOrder, logScale, logBase);
            Save();
            return this;
        }

        private enum AxisKind {
            Category,
            Value
        }

        private ExcelChart SetAxisTitle(string title, ExcelChartAxisGroup axisGroup, AxisKind axisKind) {
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            axis.RemoveAllChildren<C.Title>();
            InsertAxisTitle(axis, CreateAxisTitle(title));
            Save();
            return this;
        }

        private static void InsertAxisTitle(OpenXmlCompositeElement axis, C.Title title) {
            if (axis == null) {
                throw new ArgumentNullException(nameof(axis));
            }
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            OpenXmlElement? insertBefore = axis.GetFirstChild<C.NumberingFormat>();
            insertBefore ??= axis.GetFirstChild<C.MajorTickMark>();
            insertBefore ??= axis.GetFirstChild<C.MinorTickMark>();
            insertBefore ??= axis.GetFirstChild<C.TickLabelPosition>();
            insertBefore ??= axis.GetFirstChild<C.ShapeProperties>();
            insertBefore ??= axis.GetFirstChild<C.TextProperties>();
            insertBefore ??= axis.GetFirstChild<C.CrossingAxis>();
            insertBefore ??= axis.GetFirstChild<C.Crosses>();
            insertBefore ??= axis.GetFirstChild<C.CrossesAt>();
            insertBefore ??= axis.GetFirstChild<C.CrossBetween>();
            insertBefore ??= axis.GetFirstChild<C.MajorUnit>();
            insertBefore ??= axis.GetFirstChild<C.MinorUnit>();
            insertBefore ??= axis.GetFirstChild<C.DisplayUnits>();
            insertBefore ??= axis.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                axis.InsertBefore(title, insertBefore);
            } else {
                axis.Append(title);
            }
        }

        private ExcelChart SetAxisTitleTextStyle(ExcelChartAxisGroup axisGroup, AxisKind axisKind,
            double? fontSizePoints, bool? bold, bool? italic, string? color, string? fontName) {
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            C.Title? title = axis.GetFirstChild<C.Title>();
            if (title == null) {
                return this;
            }

            C.ChartText? chartText = title.GetFirstChild<C.ChartText>();
            if (chartText == null) {
                return this;
            }

            ApplyTextStyle(EnsureChartTextRunProperties(chartText), fontSizePoints, bold, italic, color, fontName);
            Save();
            return this;
        }

        private ExcelChart SetAxisLabelTextStyle(ExcelChartAxisGroup axisGroup, AxisKind axisKind,
            double? fontSizePoints, bool? bold, bool? italic, string? color, string? fontName) {
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            ApplyTextStyle(EnsureTextPropertiesRunProperties(axis), fontSizePoints, bold, italic, color, fontName);
            Save();
            return this;
        }

        private ExcelChart SetAxisGridlines(ExcelChartAxisGroup axisGroup, AxisKind axisKind,
            bool showMajor, bool showMinor, string? lineColor, double? lineWidthPoints) {
            ValidateAxisGridlinesStyle(lineColor, lineWidthPoints);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            ApplyGridlines(axis, showMajor, showMinor, lineColor, lineWidthPoints);
            Save();
            return this;
        }

        private ExcelChart SetAxisLabelRotation(ExcelChartAxisGroup axisGroup, AxisKind axisKind, double rotationDegrees) {
            if (double.IsNaN(rotationDegrees) || double.IsInfinity(rotationDegrees)) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees));
            }
            if (rotationDegrees < -90 || rotationDegrees > 90) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees), "Rotation must be between -90 and 90 degrees.");
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            // Ensure a complete <c:txPr> structure before mutating rotation.
            EnsureTextPropertiesRunProperties(axis);
            C.TextProperties? textProps = axis.GetFirstChild<C.TextProperties>();
            if (textProps != null) {
                A.BodyProperties body = textProps.GetFirstChild<A.BodyProperties>() ?? new A.BodyProperties();
                body.Rotation = (int)Math.Round(rotationDegrees * 60000d);
                if (body.Parent == null) {
                    textProps.Append(body);
                }
            }

            Save();
            return this;
        }

        private ExcelChart SetAxisTickLabelPosition(ExcelChartAxisGroup axisGroup, AxisKind axisKind,
            C.TickLabelPositionValues position) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            ReplaceChild(axis, new C.TickLabelPosition { Val = position });
            Save();
            return this;
        }

        private ExcelChart SetAxisNumberFormat(string formatCode, bool sourceLinked,
            ExcelChartAxisGroup axisGroup, AxisKind axisKind) {
            if (string.IsNullOrWhiteSpace(formatCode)) {
                throw new ArgumentException("Format code cannot be null or empty.", nameof(formatCode));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            C.NumberingFormat format = axis.GetFirstChild<C.NumberingFormat>() ?? new C.NumberingFormat();
            format.FormatCode = formatCode;
            format.SourceLinked = sourceLinked;
            if (format.Parent == null) {
                axis.InsertAt(format, 0);
            }

            Save();
            return this;
        }

        private static C.CategoryAxis? ResolveCategoryAxis(C.PlotArea plotArea, ExcelChartAxisGroup axisGroup) {
            var axes = plotArea.Elements<C.CategoryAxis>().ToList();
            if (axes.Count == 0) {
                return null;
            }

            bool isBar = HasHorizontalBarChart(plotArea);
            C.AxisPositionValues primaryPosition = isBar ? C.AxisPositionValues.Left : C.AxisPositionValues.Bottom;
            C.AxisPositionValues secondaryPosition = isBar ? C.AxisPositionValues.Right : C.AxisPositionValues.Top;
            C.AxisPositionValues desired = axisGroup == ExcelChartAxisGroup.Primary ? primaryPosition : secondaryPosition;

            C.CategoryAxis? axis = axes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == desired);
            if (axis != null) {
                return axis;
            }

            return axisGroup == ExcelChartAxisGroup.Primary
                ? axes.FirstOrDefault()
                : axes.Skip(1).FirstOrDefault() ?? axes.LastOrDefault();
        }

        private static C.ValueAxis? ResolveValueAxis(C.PlotArea plotArea, ExcelChartAxisGroup axisGroup) {
            var axes = plotArea.Elements<C.ValueAxis>().ToList();
            if (axes.Count == 0) {
                return null;
            }

            bool isBar = HasHorizontalBarChart(plotArea);
            C.AxisPositionValues primaryPosition = isBar ? C.AxisPositionValues.Bottom : C.AxisPositionValues.Left;
            C.AxisPositionValues secondaryPosition = isBar ? C.AxisPositionValues.Top : C.AxisPositionValues.Right;
            C.AxisPositionValues desired = axisGroup == ExcelChartAxisGroup.Primary ? primaryPosition : secondaryPosition;

            C.ValueAxis? axis = axes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == desired);
            if (axis != null) {
                return axis;
            }

            return axisGroup == ExcelChartAxisGroup.Primary
                ? axes.FirstOrDefault()
                : axes.Skip(1).FirstOrDefault() ?? axes.LastOrDefault();
        }

        private static bool HasHorizontalBarChart(C.PlotArea plotArea) {
            return plotArea.Elements<C.BarChart>()
                .Select(chart => chart.GetFirstChild<C.BarDirection>()?.Val?.Value ?? C.BarDirectionValues.Column)
                .Any(direction => direction == C.BarDirectionValues.Bar);
        }

        private static C.ValueAxis? ResolveScatterXAxis(C.PlotArea plotArea) {
            if (plotArea.Elements<C.CategoryAxis>().Any()) {
                return null;
            }

            return plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == C.AxisPositionValues.Bottom);
        }

        private static C.ValueAxis? ResolveScatterYAxis(C.PlotArea plotArea) {
            if (plotArea.Elements<C.CategoryAxis>().Any()) {
                return null;
            }

            return plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == C.AxisPositionValues.Left);
        }

    }
}
