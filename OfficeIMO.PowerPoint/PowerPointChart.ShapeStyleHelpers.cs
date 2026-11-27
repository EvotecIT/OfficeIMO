using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
        //     Represents a chart on a slide.
        public partial class PowerPointChart : PowerPointShape {
        // The fills a shape or line properties element holds at most one of.
        private static readonly string[] ShapeFills = { "solidFill", "noFill", "gradFill", "pattFill", "blipFill", "grpFill" };

        // Parents whose c:extLst the SDK reads as its generic chart extension list. Other chart elements (series, chart
        // groups, axes, data labels, the chart space) have an extension list class of their own, which a lookup of the
        // generic one never found, so positions computed from it skip theirs.
        private static readonly HashSet<string> GenericExtensionListParents = new(StringComparer.Ordinal) {
            "plotArea", "legend", "title", "scaling", "marker", "dispUnits", "trendline", "dPt", "errBars", "doughnutChart", "ofPieChart"
        };

        private void ApplyTextStyle(OfficeXmlNode runProps, OfficeLength? fontSize, bool? bold,
            bool? italic, OfficeColor? color, string? fontFamily) {
            if (fontSize != null) {
                runProps.SetInt32Attribute("sz", PowerPointTextDefaults.ToDrawingFontSize(fontSize?.ToPoints(), nameof(fontSize)));
            }
            if (bold != null) {
                runProps.SetBooleanAttribute("b", bold.Value);
            }
            if (italic != null) {
                runProps.SetBooleanAttribute("i", italic.Value);
            }
            if (color != null) {
                PowerPointTextPropertyFormatting.SetColor(runProps, color, ColorTheme);
            }
            if (fontFamily != null) {
                PowerPointTextPropertyFormatting.SetFontFamily(runProps, fontFamily);
            }
        }

        private static void ClearTextStyle(OfficeXmlNode runProps) {
            runProps.RemoveAttribute("sz");
            runProps.RemoveAttribute("b");
            runProps.RemoveAttribute("i");
            runProps.RemoveChildren(ANs, "solidFill");
            runProps.RemoveChildren(ANs, "latin");
        }

        private static void ValidateTextStyle(OfficeLength? fontSize, OfficeColor? color, string? fontFamily) {
            _ = PowerPointTextDefaults.ToDrawingFontSize(fontSize?.ToPoints(), nameof(fontSize));
            if (fontFamily != null && string.IsNullOrWhiteSpace(fontFamily)) {
                throw new ArgumentException("Font name cannot be empty.", nameof(fontFamily));
            }
        }

        private static void ValidateAxisLabelRotation(double rotationDegrees) {
            if (double.IsNaN(rotationDegrees) || double.IsInfinity(rotationDegrees)) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees));
            }
            if (rotationDegrees < -90d || rotationDegrees > 90d) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees), "Rotation must be between -90 and 90 degrees.");
            }
        }

        private static void ValidateAxisGridlinesStyle(OfficeColor? lineColor, OfficeLength? lineWidth) {
            if (lineWidth != null && lineWidth <= OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(lineWidth));
            }
        }

        private static void ValidateDataLabelLeaderLines(OfficeColor? lineColor, OfficeLength? lineWidth) {
            if (lineWidth != null && lineWidth <= OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(lineWidth));
            }
        }

        private static void ValidateAreaStyle(OfficeColor? fillColor, OfficeColor? lineColor, OfficeLength? lineWidth,
            bool noFill, bool noLine) {
            if (lineWidth != null && lineWidth <= OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(lineWidth));
            }
            if (noFill && fillColor != null) {
                throw new ArgumentException("Cannot set both fill color and noFill.", nameof(noFill));
            }
            if (noLine && (lineColor != null || lineWidth != null)) {
                throw new ArgumentException("Cannot set line color/width when noLine is true.", nameof(noLine));
            }
        }

        private static void ValidateTrendline(OfficeChartTrendlineType type, int? order, int? period,
            double? forward, double? backward, OfficeColor? lineColor, OfficeLength? lineWidth) {
            if (order != null && (order <= 0 || order > byte.MaxValue)) {
                throw new ArgumentOutOfRangeException(nameof(order));
            }
            if (period != null && period <= 0) {
                throw new ArgumentOutOfRangeException(nameof(period));
            }
            if (forward != null && forward < 0) {
                throw new ArgumentOutOfRangeException(nameof(forward));
            }
            if (backward != null && backward < 0) {
                throw new ArgumentOutOfRangeException(nameof(backward));
            }
            if (lineWidth != null && lineWidth <= OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(lineWidth));
            }

            bool isPolynomial = type == OfficeChartTrendlineType.Polynomial;
            bool isMovingAverage = type == OfficeChartTrendlineType.MovingAverage;
            if (isPolynomial && order == null) {
                throw new ArgumentException("Polynomial trendlines require an order.", nameof(order));
            }
            if (!isPolynomial && order != null) {
                throw new ArgumentException("Order is only valid for polynomial trendlines.", nameof(order));
            }
            if (isMovingAverage && period == null) {
                throw new ArgumentException("Moving average trendlines require a period.", nameof(period));
            }
            if (!isMovingAverage && period != null) {
                throw new ArgumentException("Period is only valid for moving average trendlines.", nameof(period));
            }
        }

        // ---- Element helpers ----------------------------------------------------------------------------------

        /// <summary>The first child of <paramref name="parent"/> named by <paramref name="names"/>, tried in that order.</summary>
        private static OfficeXmlNode FirstChildOf(OfficeXmlNode parent, params string[] names) {
            foreach (string name in names) {
                OfficeXmlNode child = parent.Child(CNs, name);
                if (child.Exists) return child;
            }
            return default;
        }

        /// <summary>The <c>c:extLst</c> of <paramref name="parent"/> where the SDK typed it as its generic extension list.</summary>
        private static OfficeXmlNode GenericExtensionList(OfficeXmlNode parent) =>
            GenericExtensionListParents.Contains(parent.LocalName) ? parent.Child(CNs, "extLst") : default;

        /// <summary>Inserts a detached element before <paramref name="reference"/>, or last when there is none.</summary>
        private static OfficeXmlNode InsertBeforeOrAppend(OfficeXmlNode parent, OfficeXmlNode child, OfficeXmlNode reference) =>
            reference.Exists ? parent.InsertBefore(child, reference) : parent.Append(child);

        /// <summary>Appends a new child element.</summary>
        private static OfficeXmlNode AppendNew(OfficeXmlNode parent, int namespaceId, string localName) =>
            parent.Append(parent.CreateElement(namespaceId, localName));

        /// <summary>The child element with this name, appended last when it is missing.</summary>
        private static OfficeXmlNode GetOrAppend(OfficeXmlNode parent, int namespaceId, string localName) {
            OfficeXmlNode child = parent.Child(namespaceId, localName);
            return child.Exists ? child : AppendNew(parent, namespaceId, localName);
        }

        /// <summary>
        /// Sets the <c>val</c> of a chart child element as an SDK property assignment does: an existing element in place,
        /// a new one in schema order.
        /// </summary>
        private static OfficeXmlNode SetValChild(OfficeXmlNode parent, string localName, string value, OfficeXmlChildOrder order) {
            OfficeXmlNode child = parent.Child(CNs, localName);
            if (child.IsNone) child = parent.AddOrdered(CNs, localName, order);
            child.SetAttribute("val", value);
            return child;
        }

        private static void RemoveFills(OfficeXmlNode props) {
            foreach (string fill in ShapeFills) props.RemoveChildren(ANs, fill);
        }

        private static OfficeXmlNode EnsureChartShapeProperties(OfficeXmlNode series) {
            OfficeXmlNode props = series.Child(CNs, "spPr");
            if (props.Exists) {
                props.Remove();
            } else {
                props = series.CreateElement(CNs, "spPr");
            }

            OfficeXmlNode insertAfter = FirstChildOf(series, "tx", "order", "idx");
            return insertAfter.Exists ? insertAfter.InsertAfterSelf(props) : series.InsertBefore(props, series.FirstChild);
        }

        private void ApplySolidFill(OfficeXmlNode props, OfficeColor? color) {
            RemoveFills(props);
            props.InsertOrdered(PowerPointDrawingColor.CreateFill(props, color ?? throw new ArgumentNullException(nameof(color)), ColorTheme),
                DmlShapePropertiesView.Order);
        }

        private static void ApplyNoFill(OfficeXmlNode props) {
            RemoveFills(props);
            props.AddOrdered(ANs, "noFill", DmlShapePropertiesView.Order);
        }

        private static void ApplyNoLine(OfficeXmlNode props) {
            OfficeXmlNode outline = GetOrAppend(props, ANs, "ln");
            outline.RemoveChildren(ANs, "solidFill");
            outline.RemoveChildren(ANs, "noFill");
            AppendNew(outline, ANs, "noFill");
        }

        private void ApplyLine(OfficeXmlNode props, OfficeColor? color, OfficeLength? width) {
            OfficeXmlNode outline = GetOrAppend(props, ANs, "ln");
            RemoveFills(outline);
            outline.InsertOrdered(PowerPointDrawingColor.CreateFill(outline, color ?? throw new ArgumentNullException(nameof(color)), ColorTheme),
                DmlLinePropertiesView.Order);
            if (width != null) {
                outline.SetInt32Attribute("w", checked((int)width.Value.ToEmus()));
            }
        }

        private void ApplyOptionalLine(OfficeXmlNode props, OfficeColor? color, OfficeLength? width) {
            if (color == null && width == null) {
                return;
            }

            OfficeXmlNode outline = GetOrAppend(props, ANs, "ln");
            if (color != null) {
                RemoveFills(outline);
                outline.InsertOrdered(PowerPointDrawingColor.CreateFill(outline, color.Value, ColorTheme), DmlLinePropertiesView.Order);
            }
            if (width != null) {
                outline.SetInt32Attribute("w", checked((int)width.Value.ToEmus()));
            }
        }

        private void ApplyGridlines(OfficeXmlNode axis, bool showMajor, bool showMinor,
            OfficeColor? lineColor, OfficeLength? lineWidth) {
            ApplyGridline(axis, "majorGridlines", showMajor, lineColor, lineWidth);
            ApplyGridline(axis, "minorGridlines", showMinor, lineColor, lineWidth);
        }

        private void ApplyAreaStyle(OfficeXmlNode props, OfficeColor? fillColor, OfficeColor? lineColor,
            OfficeLength? lineWidth, bool noFill, bool noLine) {
            if (noFill) {
                ApplyNoFill(props);
            } else if (fillColor != null) {
                ApplySolidFill(props, fillColor);
            }

            if (noLine) {
                ApplyNoLine(props);
            } else if (lineColor != null || lineWidth != null) {
                ApplyOptionalLine(props, lineColor, lineWidth);
            }
        }

        private void ApplyGridline(OfficeXmlNode axis, string gridlinesName, bool show,
            OfficeColor? lineColor, OfficeLength? lineWidth) {
            OfficeXmlNode gridlines = axis.Child(CNs, gridlinesName);
            if (!show) {
                if (gridlines.Exists) gridlines.Remove();
                return;
            }

            if (gridlines.IsNone) {
                gridlines = InsertAxisGridlines(axis, gridlinesName);
            }
            if (lineColor != null || lineWidth != null) {
                ApplyOptionalLine(GetOrAppend(gridlines, CNs, "spPr"), lineColor, lineWidth);
            }
        }

        private void ApplyTrendline(OfficeXmlNode series, OfficeChartTrendlineType type, int? order, int? period,
            double? forward, double? backward, double? intercept, bool displayEquation, bool displayRSquared,
            OfficeColor? lineColor, OfficeLength? lineWidth) {
            if (!IsTrendlineSupportedSeries(series)) {
                throw new InvalidOperationException("Trendlines are only supported for line, bar/column, area, and scatter series.");
            }

            series.RemoveChildren(CNs, "trendline");
            OfficeXmlNode trendline = InsertTrendline(series, series.CreateElement(CNs, "trendline"));

            if (lineColor != null || lineWidth != null) {
                ApplyOptionalLine(AppendNew(trendline, CNs, "spPr"), lineColor, lineWidth);
            }

            AppendNew(trendline, CNs, "trendlineType").SetAttribute("val", type.ToOpenXmlToken());

            if (type == OfficeChartTrendlineType.Polynomial && order != null) {
                AppendNew(trendline, CNs, "order").SetInt32Attribute("val", (byte)order.Value);
            }
            if (type == OfficeChartTrendlineType.MovingAverage && period != null) {
                AppendNew(trendline, CNs, "period").SetUInt32Attribute("val", (uint)period.Value);
            }
            if (forward != null) {
                AppendNew(trendline, CNs, "forward").SetDoubleAttribute("val", forward.Value);
            }
            if (backward != null) {
                AppendNew(trendline, CNs, "backward").SetDoubleAttribute("val", backward.Value);
            }
            if (intercept != null) {
                AppendNew(trendline, CNs, "intercept").SetDoubleAttribute("val", intercept.Value);
            }
            if (displayRSquared) {
                AppendNew(trendline, CNs, "dispRSqr").SetBooleanAttribute("val", displayRSquared);
            }
            if (displayEquation) {
                AppendNew(trendline, CNs, "dispEq").SetBooleanAttribute("val", displayEquation);
            }
        }

        private static OfficeXmlNode InsertChartSpaceShapeProperties(OfficeXmlNode chartSpace, OfficeXmlNode props) =>
            InsertBeforeOrAppend(chartSpace, props, FirstChildOf(chartSpace, "txPr", "externalData", "printSettings", "userShapes"));

        private static OfficeXmlNode InsertPlotAreaShapeProperties(OfficeXmlNode plotArea, OfficeXmlNode props) =>
            InsertBeforeOrAppend(plotArea, props, GenericExtensionList(plotArea));

        private static bool IsTrendlineSupportedSeries(OfficeXmlNode series) =>
            SeriesKind(series) is ChartSeriesKind.Line or ChartSeriesKind.Bar or ChartSeriesKind.Area or ChartSeriesKind.Scatter
                or ChartSeriesKind.Bubble;

        private static OfficeXmlNode InsertTrendline(OfficeXmlNode series, OfficeXmlNode trendline) =>
            InsertBeforeOrAppend(series, trendline, FirstChildOf(series, "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "smooth"));

        private static OfficeXmlNode InsertSeriesMarker(OfficeXmlNode series, OfficeXmlNode marker) =>
            InsertBeforeOrAppend(series, marker,
                FirstChildOf(series, "dPt", "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "smooth"));

        private static void ApplyDisplayUnitsLabel(OfficeXmlNode displayUnits, bool showLabel, string? labelText = null) {
            if (!showLabel) {
                OfficeXmlNode existing = displayUnits.Child(CNs, "dispUnitsLbl");
                if (existing.Exists) existing.Remove();
                return;
            }

            OfficeXmlNode label = GetOrAppend(displayUnits, CNs, "dispUnitsLbl");
            if (label.Child(CNs, "layout").IsNone) {
                AppendNew(label, CNs, "layout");
            }
            if (labelText != null) {
                label.RemoveChildren(CNs, "tx");
                label.Append(CreateChartText(label, labelText));
            }
        }

        private static OfficeXmlNode EnsureTextPropertiesRunProperties(OfficeXmlNode parent) {
            OfficeXmlNode textProps = parent.Child(CNs, "txPr");
            if (textProps.IsNone) {
                textProps = parent.CreateElement(CNs, "txPr");
                if (parent.Is(CNs, "catAx") || parent.Is(CNs, "valAx")) {
                    textProps = InsertBeforeOrAppend(parent, textProps, FirstChildOf(parent, "crossAx", "crosses", "crossesAt", "auto",
                        "lblAlgn", "lblOffset", "noMultiLvlLbl", "crossBetween", "majorUnit", "minorUnit", "dispUnits"));
                } else {
                    textProps = parent.Append(textProps);
                }
            }
            if (textProps.Child(ANs, "bodyPr").IsNone) {
                AppendNew(textProps, ANs, "bodyPr");
            }
            if (textProps.Child(ANs, "lstStyle").IsNone) {
                AppendNew(textProps, ANs, "lstStyle");
            }

            OfficeXmlNode paragraph = GetOrAppend(textProps, ANs, "p");
            OfficeXmlNode paragraphProps = GetOrAppend(paragraph, ANs, "pPr");
            return GetOrAppend(paragraphProps, ANs, "defRPr");
        }

        private static OfficeXmlNode InsertLegendOverlay(OfficeXmlNode legend, OfficeXmlNode overlay) {
            OfficeXmlNode insertBefore = FirstChildOf(legend, "spPr", "txPr");
            if (insertBefore.IsNone) insertBefore = GenericExtensionList(legend);
            return InsertBeforeOrAppend(legend, overlay, insertBefore);
        }

        private static OfficeXmlNode InsertAxisTitle(OfficeXmlNode axis, OfficeXmlNode title) =>
            InsertBeforeOrAppend(axis, title, FirstChildOf(axis, "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr",
                "crossAx", "crosses", "crossesAt", "auto", "lblAlgn", "lblOffset", "noMultiLvlLbl", "crossBetween", "majorUnit", "minorUnit",
                "dispUnits"));

        private static OfficeXmlNode EnsureChartTextRunProperties(OfficeXmlNode chartText) {
            OfficeXmlNode richText = GetOrAppend(chartText, CNs, "rich");
            if (richText.Child(ANs, "bodyPr").IsNone) {
                AppendNew(richText, ANs, "bodyPr");
            }
            if (richText.Child(ANs, "lstStyle").IsNone) {
                AppendNew(richText, ANs, "lstStyle");
            }

            OfficeXmlNode paragraph = GetOrAppend(richText, ANs, "p");
            OfficeXmlNode run = GetOrAppend(paragraph, ANs, "r");
            OfficeXmlNode runProps = run.Child(ANs, "rPr");
            if (runProps.IsNone) {
                runProps = run.InsertBefore(run.CreateElement(ANs, "rPr"), run.FirstChild);
            } else if (runProps != run.FirstChild) {
                runProps.Remove();
                runProps = run.InsertBefore(runProps, run.FirstChild);
            }

            return runProps;
        }

        private void ApplyMarker(OfficeXmlNode marker, OfficeChartMarkerShape style, OfficeLength? size, OfficeColor? fillColor, OfficeColor? lineColor, OfficeLength? lineWidth) {
            SetValChild(marker, "symbol", style.ToOpenXmlToken(), CMarkerView.Order);
            if (size != null) {
                byte points = checked((byte)decimal.Round(size.Value.ToEmus() / 12700m, 0, MidpointRounding.AwayFromZero));
                SetValChild(marker, "size", points.ToString(CultureInfo.InvariantCulture), CMarkerView.Order);
            }

            if (fillColor != null || lineColor != null || lineWidth != null) {
                OfficeXmlNode props = GetOrAppend(marker, CNs, "spPr");
                if (fillColor != null) {
                    ApplySolidFill(props, fillColor);
                }
                if (lineColor != null || lineWidth != null) {
                    ApplyOptionalLine(props, lineColor, lineWidth);
                }
            }
        }

    }
}
