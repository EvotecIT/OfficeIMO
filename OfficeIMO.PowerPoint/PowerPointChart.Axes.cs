using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
        //     Represents a chart on a slide.
        public partial class PowerPointChart : PowerPointShape {
        /// <summary>
        ///     Sets the category axis title.
        /// </summary>
        public PowerPointChart SetCategoryAxisTitle(string title) {
            return SetAxisTitle(CategoryAxisName, title);
        }

        /// <summary>
        ///     Sets the value axis title.
        /// </summary>
        public PowerPointChart SetValueAxisTitle(string title) {
            return SetAxisTitle(ValueAxisName, title);
        }

        /// <summary>
        ///     Sets the category axis title text style.
        /// </summary>
        public PowerPointChart SetCategoryAxisTitleTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null) {
            return SetAxisTitleTextStyle(CategoryAxisName, fontSize, bold, italic, color, fontFamily);
        }

        /// <summary>
        ///     Clears the category axis title text style.
        /// </summary>
        public PowerPointChart ClearCategoryAxisTitleTextStyle() {
            return ClearAxisTitleTextStyle(CategoryAxisName);
        }

        /// <summary>
        ///     Sets the value axis title text style.
        /// </summary>
        public PowerPointChart SetValueAxisTitleTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null) {
            return SetAxisTitleTextStyle(ValueAxisName, fontSize, bold, italic, color, fontFamily);
        }

        /// <summary>
        ///     Clears the value axis title text style.
        /// </summary>
        public PowerPointChart ClearValueAxisTitleTextStyle() {
            return ClearAxisTitleTextStyle(ValueAxisName);
        }

        /// <summary>
        ///     Sets the category axis label text style.
        /// </summary>
        public PowerPointChart SetCategoryAxisLabelTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null) {
            return SetAxisLabelTextStyle(CategoryAxisName, fontSize, bold, italic, color, fontFamily);
        }

        /// <summary>
        ///     Clears the category axis label text style.
        /// </summary>
        public PowerPointChart ClearCategoryAxisLabelTextStyle() {
            return ClearAxisLabelTextStyle(CategoryAxisName);
        }

        /// <summary>
        ///     Sets the value axis label text style.
        /// </summary>
        public PowerPointChart SetValueAxisLabelTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null) {
            return SetAxisLabelTextStyle(ValueAxisName, fontSize, bold, italic, color, fontFamily);
        }

        /// <summary>
        ///     Clears the value axis label text style.
        /// </summary>
        public PowerPointChart ClearValueAxisLabelTextStyle() {
            return ClearAxisLabelTextStyle(ValueAxisName);
        }

        /// <summary>
        ///     Sets the category axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetCategoryAxisLabelRotation(double rotationDegrees) {
            return SetAxisLabelRotation(CategoryAxisName, rotationDegrees);
        }

        /// <summary>
        ///     Sets the value axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetValueAxisLabelRotation(double rotationDegrees) {
            return SetAxisLabelRotation(ValueAxisName, rotationDegrees);
        }

        /// <summary>
        ///     Sets the category axis tick label position.
        /// </summary>
        public PowerPointChart SetCategoryAxisTickLabelPosition(OfficeChartAxisTickLabelPosition position) {
            return SetAxisTickLabelPosition(CategoryAxisName, position);
        }

        /// <summary>
        ///     Sets the value axis tick label position.
        /// </summary>
        public PowerPointChart SetValueAxisTickLabelPosition(OfficeChartAxisTickLabelPosition position) {
            return SetAxisTickLabelPosition(ValueAxisName, position);
        }

        /// <summary>
        ///     Sets category axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetCategoryAxisGridlines(bool showMajor = true, bool showMinor = false,
            OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            return SetAxisGridlines(CategoryAxisName, showMajor, showMinor, lineColor, lineWidth);
        }

        /// <summary>
        ///     Clears category axis gridlines.
        /// </summary>
        public PowerPointChart ClearCategoryAxisGridlines() {
            return ClearAxisGridlines(CategoryAxisName);
        }

        /// <summary>
        ///     Sets value axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetValueAxisGridlines(bool showMajor = true, bool showMinor = false,
            OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            return SetAxisGridlines(ValueAxisName, showMajor, showMinor, lineColor, lineWidth);
        }

        /// <summary>
        ///     Clears value axis gridlines.
        /// </summary>
        public PowerPointChart ClearValueAxisGridlines() {
            return ClearAxisGridlines(ValueAxisName);
        }

        /// <summary>
        ///     Sets scatter chart X-axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetScatterXAxisGridlines(bool showMajor = true, bool showMinor = false,
            OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisGridlines(ValueAxisName, showMajor, showMinor, lineColor, lineWidth,
                IsBottomAxis);
        }

        /// <summary>
        ///     Clears scatter chart X-axis gridlines.
        /// </summary>
        public PowerPointChart ClearScatterXAxisGridlines() {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return ClearAxisGridlines(ValueAxisName, IsBottomAxis);
        }

        /// <summary>
        ///     Sets scatter chart Y-axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetScatterYAxisGridlines(bool showMajor = true, bool showMinor = false,
            OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisGridlines(ValueAxisName, showMajor, showMinor, lineColor, lineWidth,
                IsLeftAxis);
        }

        /// <summary>
        ///     Clears scatter chart Y-axis gridlines.
        /// </summary>
        public PowerPointChart ClearScatterYAxisGridlines() {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return ClearAxisGridlines(ValueAxisName, IsLeftAxis);
        }

        /// <summary>
        ///     Sets the scatter chart X-axis title.
        /// </summary>
        public PowerPointChart SetScatterXAxisTitle(string title) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisTitle(ValueAxisName, title, IsBottomAxis);
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis title.
        /// </summary>
        public PowerPointChart SetScatterYAxisTitle(string title) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisTitle(ValueAxisName, title, IsLeftAxis);
        }

        /// <summary>
        ///     Sets the scatter chart X-axis title text style.
        /// </summary>
        public PowerPointChart SetScatterXAxisTitleTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisTitleTextStyle(ValueAxisName, fontSize, bold, italic, color, fontFamily,
                IsBottomAxis);
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis title text style.
        /// </summary>
        public PowerPointChart SetScatterYAxisTitleTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisTitleTextStyle(ValueAxisName, fontSize, bold, italic, color, fontFamily,
                IsLeftAxis);
        }

        /// <summary>
        ///     Clears the scatter chart X-axis title text style.
        /// </summary>
        public PowerPointChart ClearScatterXAxisTitleTextStyle() {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return ClearAxisTitleTextStyle(ValueAxisName, IsBottomAxis);
        }

        /// <summary>
        ///     Clears the scatter chart Y-axis title text style.
        /// </summary>
        public PowerPointChart ClearScatterYAxisTitleTextStyle() {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return ClearAxisTitleTextStyle(ValueAxisName, IsLeftAxis);
        }

        /// <summary>
        ///     Sets the scatter chart X-axis label text style.
        /// </summary>
        public PowerPointChart SetScatterXAxisLabelTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisLabelTextStyle(ValueAxisName, fontSize, bold, italic, color, fontFamily,
                IsBottomAxis);
        }

        /// <summary>
        ///     Clears the scatter chart X-axis label text style.
        /// </summary>
        public PowerPointChart ClearScatterXAxisLabelTextStyle() {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return ClearAxisLabelTextStyle(ValueAxisName, IsBottomAxis);
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis label text style.
        /// </summary>
        public PowerPointChart SetScatterYAxisLabelTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisLabelTextStyle(ValueAxisName, fontSize, bold, italic, color, fontFamily,
                IsLeftAxis);
        }

        /// <summary>
        ///     Clears the scatter chart Y-axis label text style.
        /// </summary>
        public PowerPointChart ClearScatterYAxisLabelTextStyle() {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return ClearAxisLabelTextStyle(ValueAxisName, IsLeftAxis);
        }

        /// <summary>
        ///     Sets the scatter chart X-axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetScatterXAxisLabelRotation(double rotationDegrees) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisLabelRotation(ValueAxisName, rotationDegrees, IsBottomAxis);
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetScatterYAxisLabelRotation(double rotationDegrees) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisLabelRotation(ValueAxisName, rotationDegrees, IsLeftAxis);
        }

        /// <summary>
        ///     Sets the scatter chart X-axis tick label position.
        /// </summary>
        public PowerPointChart SetScatterXAxisTickLabelPosition(OfficeChartAxisTickLabelPosition position) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisTickLabelPosition(ValueAxisName, position, IsBottomAxis);
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis tick label position.
        /// </summary>
        public PowerPointChart SetScatterYAxisTickLabelPosition(OfficeChartAxisTickLabelPosition position) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisTickLabelPosition(ValueAxisName, position, IsLeftAxis);
        }

        /// <summary>
        ///     Sets how the value axis crosses between categories.
        /// </summary>
        public PowerPointChart SetValueAxisCrossBetween(OfficeChartAxisCrossBetween between) {
            OfficeXmlNode axis = FindAxis(ValueAxisName, null);
            if (axis.IsNone) {
                return this;
            }

            ReplaceValueAxisCrossBetween(axis, between.ToOpenXmlToken());
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the category axis number format.
        /// </summary>
        public PowerPointChart SetCategoryAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetAxisNumberFormat(CategoryAxisName, formatCode, sourceLinked);
        }

        /// <summary>
        ///     Sets the value axis number format.
        /// </summary>
        public PowerPointChart SetValueAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetAxisNumberFormat(ValueAxisName, formatCode, sourceLinked);
        }

        /// <summary>
        ///     Sets the scatter chart X-axis number format.
        /// </summary>
        public PowerPointChart SetScatterXAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisNumberFormat(ValueAxisName, formatCode, sourceLinked,
                IsBottomAxis);
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis number format.
        /// </summary>
        public PowerPointChart SetScatterYAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisNumberFormat(ValueAxisName, formatCode, sourceLinked,
                IsLeftAxis);
        }

        /// <summary>
        ///     Sets display units for the scatter chart X-axis.
        /// </summary>
        public PowerPointChart SetScatterXAxisDisplayUnits(OfficeChartDisplayUnit unit, bool showLabel = true) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => SetBuiltInDisplayUnit(displayUnits, unit), showLabel, null, IsBottomAxis);
        }

        /// <summary>
        ///     Sets display units for the scatter chart X-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterXAxisDisplayUnits(OfficeChartDisplayUnit unit, string labelText, bool showLabel = true) {
            ArgumentException.ThrowIfNullOrWhiteSpace(labelText);
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => SetBuiltInDisplayUnit(displayUnits, unit), showLabel, labelText, IsBottomAxis);
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

            return SetValueAxisDisplayUnitsCore(displayUnits => SetCustomDisplayUnit(displayUnits, customUnit), showLabel, null, IsBottomAxis);
        }

        /// <summary>
        ///     Sets custom display units for the scatter chart X-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterXAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            ArgumentException.ThrowIfNullOrWhiteSpace(labelText);
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => SetCustomDisplayUnit(displayUnits, customUnit), showLabel, labelText, IsBottomAxis);
        }

        /// <summary>
        ///     Clears display units from the scatter chart X-axis.
        /// </summary>
        public PowerPointChart ClearScatterXAxisDisplayUnits() {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return ClearValueAxisDisplayUnits(IsBottomAxis);
        }

        /// <summary>
        ///     Sets display units for the scatter chart Y-axis.
        /// </summary>
        public PowerPointChart SetScatterYAxisDisplayUnits(OfficeChartDisplayUnit unit, bool showLabel = true) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => SetBuiltInDisplayUnit(displayUnits, unit), showLabel, null, IsLeftAxis);
        }

        /// <summary>
        ///     Sets display units for the scatter chart Y-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterYAxisDisplayUnits(OfficeChartDisplayUnit unit, string labelText, bool showLabel = true) {
            ArgumentException.ThrowIfNullOrWhiteSpace(labelText);
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => SetBuiltInDisplayUnit(displayUnits, unit), showLabel, labelText, IsLeftAxis);
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

            return SetValueAxisDisplayUnitsCore(displayUnits => SetCustomDisplayUnit(displayUnits, customUnit), showLabel, null, IsLeftAxis);
        }

        /// <summary>
        ///     Sets custom display units for the scatter chart Y-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterYAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            ArgumentException.ThrowIfNullOrWhiteSpace(labelText);
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => SetCustomDisplayUnit(displayUnits, customUnit), showLabel, labelText, IsLeftAxis);
        }

        /// <summary>
        ///     Clears display units from the scatter chart Y-axis.
        /// </summary>
        public PowerPointChart ClearScatterYAxisDisplayUnits() {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return ClearValueAxisDisplayUnits(IsLeftAxis);
        }

        /// <summary>
        ///     Sets display units for the value axis.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(OfficeChartDisplayUnit unit, bool showLabel = true) {
            return SetValueAxisDisplayUnitsCore(displayUnits => SetBuiltInDisplayUnit(displayUnits, unit), showLabel);
        }

        /// <summary>
        ///     Sets display units for the value axis with custom label text.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(OfficeChartDisplayUnit unit, string labelText, bool showLabel = true) {
            ArgumentException.ThrowIfNullOrWhiteSpace(labelText);

            return SetValueAxisDisplayUnitsCore(displayUnits => SetBuiltInDisplayUnit(displayUnits, unit), showLabel, labelText);
        }

        /// <summary>
        ///     Sets custom display units for the value axis.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(double customUnit, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => SetCustomDisplayUnit(displayUnits, customUnit), showLabel);
        }

        /// <summary>
        ///     Sets custom display units for the value axis with custom label text.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            ArgumentException.ThrowIfNullOrWhiteSpace(labelText);

            return SetValueAxisDisplayUnitsCore(displayUnits => SetCustomDisplayUnit(displayUnits, customUnit), showLabel, labelText);
        }

        /// <summary>
        ///     Clears display units from the value axis.
        /// </summary>
        public PowerPointChart ClearValueAxisDisplayUnits() {
            return ClearValueAxisDisplayUnits(null);
        }
    }
}
