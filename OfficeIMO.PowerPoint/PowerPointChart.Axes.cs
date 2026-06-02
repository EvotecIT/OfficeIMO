using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
        /// <summary>
        ///     Represents a chart on a slide.
        /// </summary>
        public partial class PowerPointChart : PowerPointShape {
        /// <summary>
        ///     Sets the category axis title.
        /// </summary>
        public PowerPointChart SetCategoryAxisTitle(string title) {
            return SetAxisTitle<C.CategoryAxis>(title);
        }

        /// <summary>
        ///     Sets the value axis title.
        /// </summary>
        public PowerPointChart SetValueAxisTitle(string title) {
            return SetAxisTitle<C.ValueAxis>(title);
        }

        /// <summary>
        ///     Sets the category axis title text style.
        /// </summary>
        public PowerPointChart SetCategoryAxisTitleTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            return SetAxisTitleTextStyle<C.CategoryAxis>(fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        ///     Clears the category axis title text style.
        /// </summary>
        public PowerPointChart ClearCategoryAxisTitleTextStyle() {
            return ClearAxisTitleTextStyle<C.CategoryAxis>();
        }

        /// <summary>
        ///     Sets the value axis title text style.
        /// </summary>
        public PowerPointChart SetValueAxisTitleTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            return SetAxisTitleTextStyle<C.ValueAxis>(fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        ///     Clears the value axis title text style.
        /// </summary>
        public PowerPointChart ClearValueAxisTitleTextStyle() {
            return ClearAxisTitleTextStyle<C.ValueAxis>();
        }

        /// <summary>
        ///     Sets the category axis label text style.
        /// </summary>
        public PowerPointChart SetCategoryAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            return SetAxisLabelTextStyle<C.CategoryAxis>(fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        ///     Clears the category axis label text style.
        /// </summary>
        public PowerPointChart ClearCategoryAxisLabelTextStyle() {
            return ClearAxisLabelTextStyle<C.CategoryAxis>();
        }

        /// <summary>
        ///     Sets the value axis label text style.
        /// </summary>
        public PowerPointChart SetValueAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            return SetAxisLabelTextStyle<C.ValueAxis>(fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        ///     Clears the value axis label text style.
        /// </summary>
        public PowerPointChart ClearValueAxisLabelTextStyle() {
            return ClearAxisLabelTextStyle<C.ValueAxis>();
        }

        /// <summary>
        ///     Sets the category axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetCategoryAxisLabelRotation(double rotationDegrees) {
            return SetAxisLabelRotation<C.CategoryAxis>(rotationDegrees);
        }

        /// <summary>
        ///     Sets the value axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetValueAxisLabelRotation(double rotationDegrees) {
            return SetAxisLabelRotation<C.ValueAxis>(rotationDegrees);
        }

        /// <summary>
        ///     Sets the category axis tick label position.
        /// </summary>
        public PowerPointChart SetCategoryAxisTickLabelPosition(C.TickLabelPositionValues position) {
            return SetAxisTickLabelPosition<C.CategoryAxis>(position);
        }

        /// <summary>
        ///     Sets the value axis tick label position.
        /// </summary>
        public PowerPointChart SetValueAxisTickLabelPosition(C.TickLabelPositionValues position) {
            return SetAxisTickLabelPosition<C.ValueAxis>(position);
        }

        /// <summary>
        ///     Sets category axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetCategoryAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null) {
            return SetAxisGridlines<C.CategoryAxis>(showMajor, showMinor, lineColor, lineWidthPoints);
        }

        /// <summary>
        ///     Clears category axis gridlines.
        /// </summary>
        public PowerPointChart ClearCategoryAxisGridlines() {
            return ClearAxisGridlines<C.CategoryAxis>();
        }

        /// <summary>
        ///     Sets value axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetValueAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null) {
            return SetAxisGridlines<C.ValueAxis>(showMajor, showMinor, lineColor, lineWidthPoints);
        }

        /// <summary>
        ///     Clears value axis gridlines.
        /// </summary>
        public PowerPointChart ClearValueAxisGridlines() {
            return ClearAxisGridlines<C.ValueAxis>();
        }

        /// <summary>
        ///     Sets scatter chart X-axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetScatterXAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisGridlines<C.ValueAxis>(showMajor, showMinor, lineColor, lineWidthPoints,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Clears scatter chart X-axis gridlines.
        /// </summary>
        public PowerPointChart ClearScatterXAxisGridlines() {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return ClearAxisGridlines<C.ValueAxis>(axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets scatter chart Y-axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetScatterYAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisGridlines<C.ValueAxis>(showMajor, showMinor, lineColor, lineWidthPoints,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Clears scatter chart Y-axis gridlines.
        /// </summary>
        public PowerPointChart ClearScatterYAxisGridlines() {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return ClearAxisGridlines<C.ValueAxis>(axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets the scatter chart X-axis title.
        /// </summary>
        public PowerPointChart SetScatterXAxisTitle(string title) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisTitle<C.ValueAxis>(title, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis title.
        /// </summary>
        public PowerPointChart SetScatterYAxisTitle(string title) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisTitle<C.ValueAxis>(title, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets the scatter chart X-axis title text style.
        /// </summary>
        public PowerPointChart SetScatterXAxisTitleTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisTitleTextStyle<C.ValueAxis>(fontSizePoints, bold, italic, color, fontName,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis title text style.
        /// </summary>
        public PowerPointChart SetScatterYAxisTitleTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisTitleTextStyle<C.ValueAxis>(fontSizePoints, bold, italic, color, fontName,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Clears the scatter chart X-axis title text style.
        /// </summary>
        public PowerPointChart ClearScatterXAxisTitleTextStyle() {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return ClearAxisTitleTextStyle<C.ValueAxis>(axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Clears the scatter chart Y-axis title text style.
        /// </summary>
        public PowerPointChart ClearScatterYAxisTitleTextStyle() {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return ClearAxisTitleTextStyle<C.ValueAxis>(axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets the scatter chart X-axis label text style.
        /// </summary>
        public PowerPointChart SetScatterXAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisLabelTextStyle<C.ValueAxis>(fontSizePoints, bold, italic, color, fontName,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Clears the scatter chart X-axis label text style.
        /// </summary>
        public PowerPointChart ClearScatterXAxisLabelTextStyle() {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return ClearAxisLabelTextStyle<C.ValueAxis>(axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis label text style.
        /// </summary>
        public PowerPointChart SetScatterYAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisLabelTextStyle<C.ValueAxis>(fontSizePoints, bold, italic, color, fontName,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Clears the scatter chart Y-axis label text style.
        /// </summary>
        public PowerPointChart ClearScatterYAxisLabelTextStyle() {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return ClearAxisLabelTextStyle<C.ValueAxis>(axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets the scatter chart X-axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetScatterXAxisLabelRotation(double rotationDegrees) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisLabelRotation<C.ValueAxis>(rotationDegrees, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetScatterYAxisLabelRotation(double rotationDegrees) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisLabelRotation<C.ValueAxis>(rotationDegrees, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets the scatter chart X-axis tick label position.
        /// </summary>
        public PowerPointChart SetScatterXAxisTickLabelPosition(C.TickLabelPositionValues position) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisTickLabelPosition<C.ValueAxis>(position, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis tick label position.
        /// </summary>
        public PowerPointChart SetScatterYAxisTickLabelPosition(C.TickLabelPositionValues position) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisTickLabelPosition<C.ValueAxis>(position, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets how the value axis crosses between categories.
        /// </summary>
        public PowerPointChart SetValueAxisCrossBetween(C.CrossBetweenValues between) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = plotArea.Elements<C.ValueAxis>().FirstOrDefault();
            if (axis == null) {
                return this;
            }

            ReplaceValueAxisCrossBetween(axis, new C.CrossBetween { Val = between });
            Save();
            return this;
        }

        /// <summary>
        ///     Sets the category axis number format.
        /// </summary>
        public PowerPointChart SetCategoryAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetAxisNumberFormat<C.CategoryAxis>(formatCode, sourceLinked);
        }

        /// <summary>
        ///     Sets the value axis number format.
        /// </summary>
        public PowerPointChart SetValueAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetAxisNumberFormat<C.ValueAxis>(formatCode, sourceLinked);
        }

        /// <summary>
        ///     Sets the scatter chart X-axis number format.
        /// </summary>
        public PowerPointChart SetScatterXAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisNumberFormat<C.ValueAxis>(formatCode, sourceLinked,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis number format.
        /// </summary>
        public PowerPointChart SetScatterYAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisNumberFormat<C.ValueAxis>(formatCode, sourceLinked,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets display units for the scatter chart X-axis.
        /// </summary>
        public PowerPointChart SetScatterXAxisDisplayUnits(C.BuiltInUnitValues unit, bool showLabel = true) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel, null, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets display units for the scatter chart X-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterXAxisDisplayUnits(C.BuiltInUnitValues unit, string labelText, bool showLabel = true) {
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel, labelText, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets custom display units for the scatter chart X-axis.
        /// </summary>
        public PowerPointChart SetScatterXAxisDisplayUnits(double customUnit, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel, null, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets custom display units for the scatter chart X-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterXAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel, labelText, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Clears display units from the scatter chart X-axis.
        /// </summary>
        public PowerPointChart ClearScatterXAxisDisplayUnits() {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return ClearValueAxisDisplayUnits(axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets display units for the scatter chart Y-axis.
        /// </summary>
        public PowerPointChart SetScatterYAxisDisplayUnits(C.BuiltInUnitValues unit, bool showLabel = true) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel, null, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets display units for the scatter chart Y-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterYAxisDisplayUnits(C.BuiltInUnitValues unit, string labelText, bool showLabel = true) {
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel, labelText, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets custom display units for the scatter chart Y-axis.
        /// </summary>
        public PowerPointChart SetScatterYAxisDisplayUnits(double customUnit, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel, null, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets custom display units for the scatter chart Y-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterYAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel, labelText, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Clears display units from the scatter chart Y-axis.
        /// </summary>
        public PowerPointChart ClearScatterYAxisDisplayUnits() {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return ClearValueAxisDisplayUnits(axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets display units for the value axis.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(C.BuiltInUnitValues unit, bool showLabel = true) {
            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel);
        }

        /// <summary>
        ///     Sets display units for the value axis with custom label text.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(C.BuiltInUnitValues unit, string labelText, bool showLabel = true) {
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel, labelText);
        }

        /// <summary>
        ///     Sets custom display units for the value axis.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(double customUnit, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel);
        }

        /// <summary>
        ///     Sets custom display units for the value axis with custom label text.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel, labelText);
        }

        /// <summary>
        ///     Clears display units from the value axis.
        /// </summary>
        public PowerPointChart ClearValueAxisDisplayUnits() {
            return ClearValueAxisDisplayUnits(null);
        }
    }
}
