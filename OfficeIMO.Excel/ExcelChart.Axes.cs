using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    // Represents a chart on a worksheet.
    public sealed partial class ExcelChart {
        /// <summary>
        /// Sets the category axis title.
        /// </summary>
        public ExcelChart SetCategoryAxisTitle(string title) {
            return SetCategoryAxisTitle(title, OfficeChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the category axis title for the selected axis group.
        /// </summary>
        public ExcelChart SetCategoryAxisTitle(string title, OfficeChartAxisGroup axisGroup) {
            return SetAxisTitle(title, axisGroup, AxisKind.Category);
        }

        /// <summary>
        /// Sets the value axis title.
        /// </summary>
        public ExcelChart SetValueAxisTitle(string title) {
            return SetValueAxisTitle(title, OfficeChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the value axis title for the selected axis group.
        /// </summary>
        public ExcelChart SetValueAxisTitle(string title, OfficeChartAxisGroup axisGroup) {
            return SetAxisTitle(title, axisGroup, AxisKind.Value);
        }

        /// <summary>
        /// Sets the category axis title text style.
        /// </summary>
        public ExcelChart SetCategoryAxisTitleTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisTitleTextStyle(axisGroup, AxisKind.Category, fontSize, bold, italic, color, fontFamily);
        }

        /// <summary>
        /// Sets the value axis title text style.
        /// </summary>
        public ExcelChart SetValueAxisTitleTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisTitleTextStyle(axisGroup, AxisKind.Value, fontSize, bold, italic, color, fontFamily);
        }

        /// <summary>
        /// Sets category axis gridlines visibility and optional styling.
        /// </summary>
        public ExcelChart SetCategoryAxisGridlines(bool? showMajor = null, bool? showMinor = null,
            OfficeColor? lineColor = null, OfficeLength? lineWidth = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisGridlines(axisGroup, AxisKind.Category, showMajor, showMinor, lineColor, lineWidth);
        }

        /// <summary>
        /// Sets value axis gridlines visibility and optional styling.
        /// </summary>
        public ExcelChart SetValueAxisGridlines(bool? showMajor = null, bool? showMinor = null,
            OfficeColor? lineColor = null, OfficeLength? lineWidth = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisGridlines(axisGroup, AxisKind.Value, showMajor, showMinor, lineColor, lineWidth);
        }

        /// <summary>
        /// Sets category axis line visibility and optional styling.
        /// </summary>
        public ExcelChart SetCategoryAxisLine(OfficeColor? lineColor = null, OfficeLength? lineWidth = null,
            bool noLine = false, OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisLine(axisGroup, AxisKind.Category, lineColor, lineWidth, noLine);
        }

        /// <summary>
        /// Sets value axis line visibility and optional styling.
        /// </summary>
        public ExcelChart SetValueAxisLine(OfficeColor? lineColor = null, OfficeLength? lineWidth = null,
            bool noLine = false, OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisLine(axisGroup, AxisKind.Value, lineColor, lineWidth, noLine);
        }

        /// <summary>
        /// Sets the category axis label text style.
        /// </summary>
        public ExcelChart SetCategoryAxisLabelTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisLabelTextStyle(axisGroup, AxisKind.Category, fontSize, bold, italic, color, fontFamily);
        }

        /// <summary>
        /// Sets the value axis label text style.
        /// </summary>
        public ExcelChart SetValueAxisLabelTextStyle(OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisLabelTextStyle(axisGroup, AxisKind.Value, fontSize, bold, italic, color, fontFamily);
        }

        /// <summary>
        /// Sets the category axis label rotation in degrees (-90..90).
        /// </summary>
        public ExcelChart SetCategoryAxisLabelRotation(double rotationDegrees,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisLabelRotation(axisGroup, AxisKind.Category, rotationDegrees);
        }

        /// <summary>
        /// Sets the value axis label rotation in degrees (-90..90).
        /// </summary>
        public ExcelChart SetValueAxisLabelRotation(double rotationDegrees,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisLabelRotation(axisGroup, AxisKind.Value, rotationDegrees);
        }

        /// <summary>
        /// Sets the category axis tick label position.
        /// </summary>
        public ExcelChart SetCategoryAxisTickLabelPosition(OfficeChartAxisTickLabelPosition position,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisTickLabelPosition(axisGroup, AxisKind.Category, position);
        }

        /// <summary>
        /// Sets the value axis tick label position.
        /// </summary>
        public ExcelChart SetValueAxisTickLabelPosition(OfficeChartAxisTickLabelPosition position,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            return SetAxisTickLabelPosition(axisGroup, AxisKind.Value, position);
        }

        /// <summary>
        /// Sets how the value axis crosses between categories.
        /// </summary>
        public ExcelChart SetValueAxisCrossBetween(OfficeChartAxisCrossBetween between,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis.IsNone) {
                return this;
            }

            ReplaceChild(axis, axis.NewVal(C, "crossBetween", between.ToOpenXmlToken()));
            Save();
            return this;
        }

        /// <summary>
        /// Sets where the category axis crosses the value axis.
        /// </summary>
        public ExcelChart SetCategoryAxisCrossing(OfficeChartAxisCrossingPosition crosses, double? crossesAt = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveCategoryAxis(plotArea, axisGroup);
            if (axis.IsNone) {
                return this;
            }

            ApplyAxisCrossing(axis, crosses, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        /// Sets where the value axis crosses the category axis.
        /// </summary>
        public ExcelChart SetValueAxisCrossing(OfficeChartAxisCrossingPosition crosses, double? crossesAt = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis.IsNone) {
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
        public ExcelChart SetScatterXAxisCrossing(OfficeChartAxisCrossingPosition crosses, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveScatterXAxis(plotArea);
            if (axis.IsNone) {
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
        public ExcelChart SetScatterYAxisCrossing(OfficeChartAxisCrossingPosition crosses, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveScatterYAxis(plotArea);
            if (axis.IsNone) {
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
        public ExcelChart SetValueAxisDisplayUnits(OfficeChartDisplayUnit unit, bool showLabel = true,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode existingDisplayUnits = axis.Child(C, "dispUnits");
            OfficeXmlNode displayUnits = existingDisplayUnits.Exists ? existingDisplayUnits : axis.New(C, "dispUnits");
            displayUnits.RemoveChildren(C, "custUnit");
            displayUnits.RemoveChildren(C, "builtInUnit");
            displayUnits.Append(displayUnits.NewVal(C, "builtInUnit", unit.ToOpenXmlToken()));
            ApplyDisplayUnitsLabel(displayUnits, showLabel);
            if (existingDisplayUnits.IsNone) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets display units for the value axis with custom label text.
        /// </summary>
        public ExcelChart SetValueAxisDisplayUnits(OfficeChartDisplayUnit unit, string labelText, bool showLabel = true,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            ArgumentException.ThrowIfNullOrWhiteSpace(labelText);

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode existingDisplayUnits = axis.Child(C, "dispUnits");
            OfficeXmlNode displayUnits = existingDisplayUnits.Exists ? existingDisplayUnits : axis.New(C, "dispUnits");
            displayUnits.RemoveChildren(C, "custUnit");
            displayUnits.RemoveChildren(C, "builtInUnit");
            displayUnits.Append(displayUnits.NewVal(C, "builtInUnit", unit.ToOpenXmlToken()));
            ApplyDisplayUnitsLabel(displayUnits, showLabel, labelText);
            if (existingDisplayUnits.IsNone) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets custom display units for the value axis.
        /// </summary>
        public ExcelChart SetValueAxisDisplayUnits(double customUnit, bool showLabel = true,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            if (customUnit <= 0 || double.IsNaN(customUnit) || double.IsInfinity(customUnit)) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode existingDisplayUnits = axis.Child(C, "dispUnits");
            OfficeXmlNode displayUnits = existingDisplayUnits.Exists ? existingDisplayUnits : axis.New(C, "dispUnits");
            displayUnits.RemoveChildren(C, "builtInUnit");
            displayUnits.RemoveChildren(C, "custUnit");
            displayUnits.Append(displayUnits.NewVal(C, "custUnit", customUnit));
            ApplyDisplayUnitsLabel(displayUnits, showLabel);
            if (existingDisplayUnits.IsNone) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets custom display units for the value axis with custom label text.
        /// </summary>
        public ExcelChart SetValueAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            if (customUnit <= 0 || double.IsNaN(customUnit) || double.IsInfinity(customUnit)) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            ArgumentException.ThrowIfNullOrWhiteSpace(labelText);

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode existingDisplayUnits = axis.Child(C, "dispUnits");
            OfficeXmlNode displayUnits = existingDisplayUnits.Exists ? existingDisplayUnits : axis.New(C, "dispUnits");
            displayUnits.RemoveChildren(C, "builtInUnit");
            displayUnits.RemoveChildren(C, "custUnit");
            displayUnits.Append(displayUnits.NewVal(C, "custUnit", customUnit));
            ApplyDisplayUnitsLabel(displayUnits, showLabel, labelText);
            if (existingDisplayUnits.IsNone) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Clears display units from the value axis.
        /// </summary>
        public ExcelChart ClearValueAxisDisplayUnits(OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis.IsNone) {
                return this;
            }

            axis.RemoveChild(C, "dispUnits");
            Save();
            return this;
        }

        /// <summary>
        /// Sets the category axis orientation (normal or reversed order).
        /// </summary>
        public ExcelChart SetCategoryAxisReverseOrder(bool reverseOrder = true,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveCategoryAxis(plotArea, axisGroup);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode scaling = EnsureScaling(axis);
            ReplaceChild(scaling, scaling.NewVal(C, "orientation", reverseOrder ? "maxMin" : "minMax"));
            Save();
            return this;
        }

        /// <summary>
        /// Sets category axis scale parameters for line, bar, column, and area charts.
        /// </summary>
        /// <param name="minimum">Optional minimum category-axis value.</param>
        /// <param name="maximum">Optional maximum category-axis value.</param>
        /// <param name="majorUnit">Optional major unit spacing.</param>
        /// <param name="minorUnit">Optional minor unit spacing.</param>
        /// <param name="reverseOrder">Optional category order direction.</param>
        /// <param name="axisGroup">Primary or secondary category axis.</param>
        public ExcelChart SetCategoryAxisScale(double? minimum = null, double? maximum = null,
            double? majorUnit = null, double? minorUnit = null, bool? reverseOrder = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale: null, logBase: null);

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveDateAxis(plotArea, axisGroup);
            if (axis.IsNone) {
                OfficeXmlNode categoryAxis = ResolveCategoryAxis(plotArea, axisGroup);
                if (categoryAxis.IsNone) {
                    return this;
                }

                axis = minimum != null || maximum != null || majorUnit != null || minorUnit != null
                    ? PromoteCategoryAxisToDateAxis(plotArea, categoryAxis)
                    : categoryAxis;
            }

            ApplyAxisScale(axis, minimum, maximum, majorUnit, minorUnit, reverseOrder, logScale: null, logBase: null);
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

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveScatterXAxis(plotArea);
            if (axis.IsNone) {
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

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveScatterYAxis(plotArea);
            if (axis.IsNone) {
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
            return SetCategoryAxisNumberFormat(formatCode, sourceLinked, OfficeChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the category axis number format for the selected axis group.
        /// </summary>
        public ExcelChart SetCategoryAxisNumberFormat(string formatCode, bool sourceLinked, OfficeChartAxisGroup axisGroup) {
            return SetAxisNumberFormat(formatCode, sourceLinked, axisGroup, AxisKind.Category);
        }

        /// <summary>
        /// Sets the value axis number format.
        /// </summary>
        public ExcelChart SetValueAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetValueAxisNumberFormat(formatCode, sourceLinked, OfficeChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the value axis number format for the selected axis group.
        /// </summary>
        public ExcelChart SetValueAxisNumberFormat(string formatCode, bool sourceLinked, OfficeChartAxisGroup axisGroup) {
            return SetAxisNumberFormat(formatCode, sourceLinked, axisGroup, AxisKind.Value);
        }

        /// <summary>
        /// Sets value axis scale parameters for the selected axis group.
        /// </summary>
        public ExcelChart SetValueAxisScale(double? minimum = null, double? maximum = null, double? majorUnit = null,
            double? minorUnit = null, double? logBase = null, bool? reverseOrder = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary, bool? logScale = null) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale, logBase);

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis.IsNone) {
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

        /// <summary>
        /// The category or value axis of a group; the primary category axis of a scatter chart is its bottom value axis.
        /// </summary>
        private static OfficeXmlNode ResolveAxis(OfficeXmlNode plotArea, OfficeChartAxisGroup axisGroup, AxisKind axisKind) {
            OfficeXmlNode axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis.IsNone && axisKind == AxisKind.Category && axisGroup == OfficeChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            return axis;
        }

        private ExcelChart SetAxisTitle(string title, OfficeChartAxisGroup axisGroup, AxisKind axisKind) {
            ArgumentNullException.ThrowIfNull(title);

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveAxis(plotArea, axisGroup, axisKind);
            if (axis.IsNone) {
                return this;
            }

            axis.RemoveChildren(C, "title");
            InsertAxisTitle(axis, CreateAxisTitle(axis, title));
            Save();
            return this;
        }

        private static void InsertAxisTitle(OfficeXmlNode axis, OfficeXmlNode title) {
            if (axis.IsNone) throw new ArgumentNullException(nameof(axis));
            if (title.IsNone) throw new ArgumentNullException(nameof(title));

            OfficeXmlNode insertBefore = axis.Child(C, "numFmt");
            if (insertBefore.IsNone) insertBefore = axis.Child(C, "majorTickMark");
            if (insertBefore.IsNone) insertBefore = axis.Child(C, "minorTickMark");
            if (insertBefore.IsNone) insertBefore = axis.Child(C, "tickLblPos");
            if (insertBefore.IsNone) insertBefore = AxisShapePropertiesQuirk.AuthoredShapeProperties(axis);
            if (insertBefore.IsNone) insertBefore = FindAxisChildFromTextProperties(axis);

            if (insertBefore.Exists) {
                axis.InsertBefore(title, insertBefore);
            } else {
                axis.Append(title);
            }
        }

        private static void InsertAxisShapeProperties(OfficeXmlNode axis, OfficeXmlNode properties) {
            OfficeXmlNode insertBefore = FindAxisChildFromTextProperties(axis);
            AxisShapePropertiesQuirk.InsertAuthoredShapeProperties(axis, properties, insertBefore);
        }

        /// <summary>The first of the axis children from <c>c:txPr</c> on that an axis title or shape properties go before.</summary>
        private static OfficeXmlNode FindAxisChildFromTextProperties(OfficeXmlNode axis) {
            OfficeXmlNode insertBefore = axis.Child(C, "txPr");
            if (insertBefore.IsNone) insertBefore = axis.Child(C, "crossAx");
            if (insertBefore.IsNone) insertBefore = axis.Child(C, "crosses");
            if (insertBefore.IsNone) insertBefore = axis.Child(C, "crossesAt");
            if (insertBefore.IsNone) insertBefore = axis.Child(C, "crossBetween");
            if (insertBefore.IsNone) insertBefore = axis.Child(C, "majorUnit");
            if (insertBefore.IsNone) insertBefore = axis.Child(C, "minorUnit");
            if (insertBefore.IsNone) insertBefore = axis.Child(C, "dispUnits");
            if (insertBefore.IsNone) insertBefore = AxisShapePropertiesQuirk.GenericExtensionList(axis);
            return insertBefore;
        }

        private ExcelChart SetAxisTitleTextStyle(OfficeChartAxisGroup axisGroup, AxisKind axisKind,
            OfficeLength? fontSize, bool? bold, bool? italic, OfficeColor? color, string? fontFamily) {
            ValidateDataLabelTextStyle(fontSize, color, fontFamily);

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveAxis(plotArea, axisGroup, axisKind);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode title = axis.Child(C, "title");
            if (title.IsNone) {
                return this;
            }

            OfficeXmlNode chartText = title.Child(C, "tx");
            if (chartText.IsNone) {
                return this;
            }

            ApplyTextStyle(EnsureChartTextRunProperties(chartText), fontSize, bold, italic, color, fontFamily);
            Save();
            return this;
        }

        private ExcelChart SetAxisLabelTextStyle(OfficeChartAxisGroup axisGroup, AxisKind axisKind,
            OfficeLength? fontSize, bool? bold, bool? italic, OfficeColor? color, string? fontFamily) {
            ValidateDataLabelTextStyle(fontSize, color, fontFamily);

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveAxis(plotArea, axisGroup, axisKind);
            if (axis.IsNone) {
                return this;
            }

            ApplyTextStyle(EnsureTextPropertiesRunProperties(axis), fontSize, bold, italic, color, fontFamily);
            Save();
            return this;
        }

        private ExcelChart SetAxisGridlines(OfficeChartAxisGroup axisGroup, AxisKind axisKind,
            bool? showMajor, bool? showMinor, OfficeColor? lineColor, OfficeLength? lineWidth) {
            ValidateAxisGridlinesStyle(lineColor, lineWidth);

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveAxis(plotArea, axisGroup, axisKind);
            if (axis.IsNone) {
                return this;
            }

            ApplyGridlines(axis, showMajor, showMinor, lineColor, lineWidth);
            Save();
            return this;
        }

        private ExcelChart SetAxisLine(OfficeChartAxisGroup axisGroup, AxisKind axisKind,
            OfficeColor? lineColor, OfficeLength? lineWidth, bool noLine) {
            ValidateAxisLineStyle(lineColor, lineWidth, noLine);

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveAxis(plotArea, axisGroup, axisKind);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode existing = AxisShapePropertiesQuirk.AuthoredShapeProperties(axis);
            OfficeXmlNode props = existing.Exists ? existing : AxisShapePropertiesQuirk.NewAuthoredShapeProperties();
            if (noLine) {
                ApplyNoLine(props);
            } else if (lineColor != null || lineWidth != null) {
                OfficeColor? normalizedLine = lineColor != null ? lineColor : null;
                ApplyOptionalLine(props, normalizedLine, lineWidth);
            } else {
                return this;
            }

            if (existing.IsNone) {
                InsertAxisShapeProperties(axis, props);
            }

            Save();
            return this;
        }

        private static void ValidateAxisLineStyle(OfficeColor? lineColor, OfficeLength? lineWidth, bool noLine) {
            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth));
            if (lineWidth != null && lineWidth.Value <= OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(lineWidth));
            }
            if (noLine && (lineColor != null || lineWidth != null)) {
                throw new ArgumentException("Cannot set axis line color/width when noLine is true.", nameof(noLine));
            }
        }

        private ExcelChart SetAxisLabelRotation(OfficeChartAxisGroup axisGroup, AxisKind axisKind, double rotationDegrees) {
            if (double.IsNaN(rotationDegrees) || double.IsInfinity(rotationDegrees)) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees));
            }
            if (rotationDegrees < -90 || rotationDegrees > 90) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees), "Rotation must be between -90 and 90 degrees.");
            }

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveAxis(plotArea, axisGroup, axisKind);
            if (axis.IsNone) {
                return this;
            }

            // Ensure a complete <c:txPr> structure before mutating rotation.
            EnsureTextPropertiesRunProperties(axis);
            OfficeXmlNode textProps = axis.Child(C, "txPr");
            if (textProps.Exists) {
                OfficeXmlNode body = textProps.Child(A, "bodyPr");
                if (body.IsNone) body = textProps.Append(textProps.New(A, "bodyPr"));
                body.SetInt32Attribute("rot", (int)Math.Round(rotationDegrees * 60000d));
            }

            Save();
            return this;
        }

        private ExcelChart SetAxisTickLabelPosition(OfficeChartAxisGroup axisGroup, AxisKind axisKind,
            OfficeChartAxisTickLabelPosition position) {
            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveAxis(plotArea, axisGroup, axisKind);
            if (axis.IsNone) {
                return this;
            }

            ReplaceChild(axis, axis.NewVal(C, "tickLblPos", position.ToOpenXmlToken()));
            Save();
            return this;
        }

        private ExcelChart SetAxisNumberFormat(string formatCode, bool sourceLinked,
            OfficeChartAxisGroup axisGroup, AxisKind axisKind) {
            ArgumentException.ThrowIfNullOrWhiteSpace(formatCode);

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveAxis(plotArea, axisGroup, axisKind);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode existing = axis.Child(C, "numFmt");
            OfficeXmlNode format = existing.Exists ? existing : axis.New(C, "numFmt");
            format.SetAttribute("formatCode", formatCode);
            format.SetBooleanAttribute("sourceLinked", sourceLinked);
            if (existing.IsNone) {
                axis.InsertBefore(format, axis.FirstChild);
            }

            Save();
            return this;
        }

        private static OfficeXmlNode ResolveCategoryAxis(OfficeXmlNode plotArea, OfficeChartAxisGroup axisGroup) =>
            ResolveAxisByPosition(plotArea, "catAx", axisGroup, valueAxis: false);

        private static OfficeXmlNode ResolveDateAxis(OfficeXmlNode plotArea, OfficeChartAxisGroup axisGroup) =>
            ResolveAxisByPosition(plotArea, "dateAx", axisGroup, valueAxis: false);

        private static OfficeXmlNode ResolveValueAxis(OfficeXmlNode plotArea, OfficeChartAxisGroup axisGroup) =>
            ResolveAxisByPosition(plotArea, "valAx", axisGroup, valueAxis: true);

        /// <summary>
        /// The axis of this kind at the group's position (bottom or left for the primary group, swapped for horizontal
        /// bars), else the first axis for the primary group or the second (or last) for the secondary group.
        /// </summary>
        private static OfficeXmlNode ResolveAxisByPosition(OfficeXmlNode plotArea, string axisName, OfficeChartAxisGroup axisGroup, bool valueAxis) {
            List<OfficeXmlNode> axes = plotArea.Children(C, axisName).ToList();
            if (axes.Count == 0) {
                return default;
            }

            bool isBar = HasHorizontalBarChart(plotArea);
            string primaryPosition = valueAxis ? (isBar ? "b" : "l") : (isBar ? "l" : "b");
            string secondaryPosition = valueAxis ? (isBar ? "t" : "r") : (isBar ? "r" : "t");
            string desired = axisGroup == OfficeChartAxisGroup.Primary ? primaryPosition : secondaryPosition;

            foreach (OfficeXmlNode axis in axes) {
                if (axis.ChildVal(C, "axPos") == desired) {
                    return axis;
                }
            }

            return axisGroup == OfficeChartAxisGroup.Primary
                ? axes[0]
                : axes.Count > 1 ? axes[1] : axes[axes.Count - 1];
        }

        private static OfficeXmlNode PromoteCategoryAxisToDateAxis(OfficeXmlNode plotArea, OfficeXmlNode categoryAxis) {
            OfficeXmlNode dateAxis = plotArea.New(C, "dateAx");
            AppendClone(dateAxis, categoryAxis.Child(C, "axId"));
            AppendClone(dateAxis, categoryAxis.Child(C, "scaling"));
            AppendClone(dateAxis, categoryAxis.Child(C, "delete"));
            AppendClone(dateAxis, categoryAxis.Child(C, "axPos"));
            AppendClone(dateAxis, categoryAxis.Child(C, "majorGridlines"));
            AppendClone(dateAxis, categoryAxis.Child(C, "minorGridlines"));
            AppendClone(dateAxis, categoryAxis.Child(C, "title"));
            AppendClone(dateAxis, categoryAxis.Child(C, "numFmt"));
            AppendClone(dateAxis, categoryAxis.Child(C, "majorTickMark"));
            AppendClone(dateAxis, categoryAxis.Child(C, "minorTickMark"));
            AppendClone(dateAxis, categoryAxis.Child(C, "tickLblPos"));
            AppendClone(dateAxis, AxisShapePropertiesQuirk.ChartShapeProperties(categoryAxis));
            AppendClone(dateAxis, categoryAxis.Child(C, "txPr"));
            AppendClone(dateAxis, categoryAxis.Child(C, "crossAx"));
            AppendClone(dateAxis, categoryAxis.Child(C, "crosses"));
            AppendClone(dateAxis, categoryAxis.Child(C, "crossesAt"));
            AppendClone(dateAxis, categoryAxis.Child(C, "lblOffset"));

            return categoryAxis.Parent.Exists
                ? categoryAxis.ReplaceWith(dateAxis)
                : plotArea.Append(dateAxis);
        }

        private static void AppendClone(OfficeXmlNode destination, OfficeXmlNode child) {
            if (child.Exists) {
                destination.Append(child.Clone());
            }
        }

        private static bool HasHorizontalBarChart(OfficeXmlNode plotArea) {
            foreach (OfficeXmlNode chart in plotArea.Children(C, "barChart")) {
                if ((chart.ChildVal(C, "barDir") ?? "col") == "bar") return true;
            }
            return false;
        }

        private static OfficeXmlNode ResolveScatterXAxis(OfficeXmlNode plotArea) => ResolveScatterAxis(plotArea, "b");

        private static OfficeXmlNode ResolveScatterYAxis(OfficeXmlNode plotArea) => ResolveScatterAxis(plotArea, "l");

        /// <summary>The first value axis at this position when the plot area has no category axis, or none.</summary>
        private static OfficeXmlNode ResolveScatterAxis(OfficeXmlNode plotArea, string position) {
            if (plotArea.Child(C, "catAx").Exists) {
                return default;
            }

            foreach (OfficeXmlNode axis in plotArea.Children(C, "valAx")) {
                if (axis.ChildVal(C, "axPos") == position) return axis;
            }
            return default;
        }

    }
}
