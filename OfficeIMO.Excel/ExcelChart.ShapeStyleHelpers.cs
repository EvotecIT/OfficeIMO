using System.Globalization;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    // Represents a chart on a worksheet.
    public sealed partial class ExcelChart {
        // Elements that follow c:trendline in a series, in the order the insertion point is looked up.
        private static readonly string[] SeriesChildrenAfterTrendline = {
            "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "smooth", "extLst"
        };

        // Elements that follow c:txPr in a category or value axis, in the order the insertion point is looked up.
        private static readonly string[] AxisChildrenAfterTextProperties = {
            "crossAx", "crosses", "crossesAt", "auto", "lblAlgn", "lblOffset", "noMultiLvlLbl", "crossBetween",
            "majorUnit", "minorUnit", "dispUnits", "extLst"
        };

        /// <summary>The first existing chart child among <paramref name="names"/>, looked up in that order, or none.</summary>
        private static OfficeXmlNode FirstExistingChild(OfficeXmlNode parent, string[] names) {
            foreach (string name in names) {
                OfficeXmlNode child = parent.Child(C, name);
                if (child.Exists) return child;
            }
            return OfficeXmlNode.None;
        }

        private static void ApplyGridlines(OfficeXmlNode axis, bool? showMajor, bool? showMinor,
            OfficeColor? lineColor, OfficeLength? lineWidth) {
            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth), allowZero: true);
            OfficeXmlNode existingMajor = axis.Child(C, "majorGridlines");
            if (showMajor == true || showMajor == null && existingMajor.Exists) {
                OfficeXmlNode major = existingMajor.Exists ? existingMajor : axis.Append(axis.New(C, "majorGridlines"));
                ApplyGridlineStyle(major, lineColor, lineWidth);
            } else if (existingMajor.Exists) {
                existingMajor.Remove();
            }

            OfficeXmlNode existingMinor = axis.Child(C, "minorGridlines");
            if (showMinor == true || showMinor == null && existingMinor.Exists) {
                OfficeXmlNode minor = existingMinor.Exists ? existingMinor : axis.Append(axis.New(C, "minorGridlines"));
                ApplyGridlineStyle(minor, lineColor, lineWidth);
            } else if (existingMinor.Exists) {
                existingMinor.Remove();
            }
        }

        private static void ApplyTrendline(OfficeXmlNode series, OfficeChartTrendlineType type, int? order, int? period,
            double? forward, double? backward, double? intercept, bool displayEquation, bool displayRSquared,
            OfficeColor? lineColor, OfficeLength? lineWidth) {
            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth), allowZero: true);
            if (!IsTrendlineSupportedSeries(series)) {
                throw new InvalidOperationException("Trendlines are only supported for line, bar/column, area, scatter, and bubble series.");
            }

            series.RemoveChildren(C, "trendline");
            // The trendline goes in place empty and is filled there; its position doesn't depend on its content.
            OfficeXmlNode trendline = InsertTrendline(series, series.New(C, "trendline"));

            if (lineColor != null || lineWidth != null) {
                OfficeXmlNode props = trendline.Append(trendline.New(C, "spPr"));
                OfficeColor? normalizedLine = lineColor != null ? lineColor : null;
                ApplyOptionalLine(props, normalizedLine, lineWidth);
            }

            trendline.Append(trendline.NewVal(C, "trendlineType", type.ToOpenXmlToken()));

            if (type == OfficeChartTrendlineType.Polynomial && order != null) {
                trendline.Append(trendline.NewVal(C, "order", (int)(byte)order.Value));
            }
            if (type == OfficeChartTrendlineType.MovingAverage && period != null) {
                trendline.Append(trendline.NewVal(C, "period", (uint)period.Value));
            }
            if (forward != null) {
                trendline.Append(trendline.NewVal(C, "forward", forward.Value));
            }
            if (backward != null) {
                trendline.Append(trendline.NewVal(C, "backward", backward.Value));
            }
            if (intercept != null) {
                trendline.Append(trendline.NewVal(C, "intercept", intercept.Value));
            }

            if (displayRSquared) {
                trendline.Append(trendline.NewVal(C, "dispRSqr", displayRSquared));
            }
            if (displayEquation) {
                trendline.Append(trendline.NewVal(C, "dispEq", displayEquation));
            }
        }

        private static bool IsTrendlineSupportedSeries(OfficeXmlNode series) {
            return IsLineSeries(series)
                || IsSeriesUnder(series, "barChart", "bar3DChart")
                || IsSeriesUnder(series, "areaChart", "area3DChart")
                || IsSeriesUnder(series, "scatterChart")
                || IsSeriesUnder(series, "bubbleChart");
        }

        private static OfficeXmlNode InsertTrendline(OfficeXmlNode series, OfficeXmlNode trendline) {
            OfficeXmlNode insertBefore = FirstExistingChild(series, SeriesChildrenAfterTrendline);
            return series.InsertBefore(trendline, insertBefore);
        }

        private static void ApplyGridlineStyle(OfficeXmlNode gridlines, OfficeColor? lineColor, OfficeLength? lineWidth) {
            if (lineColor == null && lineWidth == null) {
                return;
            }
            OfficeXmlNode existing = gridlines.Child(C, "spPr");
            OfficeXmlNode props = existing.Exists ? existing : gridlines.Append(gridlines.New(C, "spPr"));
            OfficeColor? normalizedLine = lineColor != null ? lineColor : null;
            ApplyOptionalLine(props, normalizedLine, lineWidth);
        }

        private static void ApplyAxisCrossing(OfficeXmlNode axis, OfficeChartAxisCrossingPosition crosses, double? crossesAt) {
            axis.RemoveChild(C, "crosses");
            axis.RemoveChild(C, "crossesAt");

            OfficeXmlNode crossing = crossesAt != null
                ? axis.NewVal(C, "crossesAt", crossesAt.Value)
                : axis.NewVal(C, "crosses", crosses.ToOpenXmlToken());

            OfficeXmlNode crossAxis = axis.Child(C, "crossAx");
            if (crossAxis.Exists) {
                crossAxis.InsertAfterSelf(crossing);
            } else {
                axis.Append(crossing);
            }
        }

        private static void ApplyDisplayUnitsLabel(OfficeXmlNode displayUnits, bool showLabel, string? labelText = null) {
            OfficeXmlNode existing = displayUnits.Child(C, "dispUnitsLbl");
            if (!showLabel) {
                if (existing.Exists) existing.Remove();
                return;
            }

            OfficeXmlNode label = existing.Exists ? existing : displayUnits.Append(displayUnits.New(C, "dispUnitsLbl"));
            if (label.Child(C, "layout").IsNone) {
                label.Append(label.New(C, "layout"));
            }
            if (labelText != null) {
                label.RemoveChildren(C, "tx");
                label.Append(CreateChartText(label, labelText));
            }
        }

        private static OfficeXmlNode EnsureDataLabelShapeProperties(OfficeXmlNode labels) {
            OfficeXmlNode props = labels.Child(C, "spPr");
            return props.Exists ? props : labels.Append(labels.New(C, "spPr"));
        }

        private static void ApplyNoFill(OfficeXmlNode props) {
            RemoveShapeFillChoices(props);
            InsertShapeFill(props, props.New(A, "noFill"));
        }

        private static void ApplyNoLine(OfficeXmlNode props) {
            OfficeXmlNode existing = props.Child(A, "ln");
            OfficeXmlNode outline = existing.Exists ? existing : props.New(A, "ln");
            RemoveShapeFillChoices(outline);
            outline.InsertBefore(outline.New(A, "noFill"), outline.FirstChild);
            if (existing.IsNone) {
                InsertOutline(props, outline);
            }
        }

        private static OfficeXmlNode EnsureTextPropertiesRunProperties(OfficeXmlNode parent) {
            OfficeXmlNode textProps = parent.Child(C, "txPr");
            if (textProps.IsNone) {
                // The text properties go in place empty and are filled there; where they go doesn't depend on their
                // content.
                textProps = parent.New(C, "txPr");
                if (parent.Is(C, "catAx") || parent.Is(C, "valAx")) {
                    OfficeXmlNode insertBefore = FirstExistingChild(parent, AxisChildrenAfterTextProperties);
                    textProps = parent.InsertBefore(textProps, insertBefore);
                } else {
                    textProps = parent.Append(textProps);
                }
            }

            return FillTextPropertiesRunProperties(textProps);
        }

        // Adds what is missing of c:txPr/a:bodyPr, a:lstStyle and a:p/a:pPr/a:defRPr and returns the a:defRPr.
        private static OfficeXmlNode FillTextPropertiesRunProperties(OfficeXmlNode textProps) {
            if (textProps.Child(A, "bodyPr").IsNone) {
                textProps.Append(textProps.New(A, "bodyPr"));
            }
            if (textProps.Child(A, "lstStyle").IsNone) {
                textProps.Append(textProps.New(A, "lstStyle"));
            }

            OfficeXmlNode paragraph = textProps.Child(A, "p");
            if (paragraph.IsNone) {
                paragraph = textProps.Append(textProps.New(A, "p"));
            }

            OfficeXmlNode paragraphProps = paragraph.Child(A, "pPr");
            if (paragraphProps.IsNone) {
                paragraphProps = paragraph.Append(paragraph.New(A, "pPr"));
            }

            OfficeXmlNode runProps = paragraphProps.Child(A, "defRPr");
            if (runProps.IsNone) {
                runProps = paragraphProps.Append(paragraphProps.New(A, "defRPr"));
            }

            return runProps;
        }

        private static OfficeXmlNode EnsureChartTextRunProperties(OfficeXmlNode chartText) {
            OfficeXmlNode richText = chartText.Child(C, "rich");
            if (richText.IsNone) {
                // A new c:rich goes last, as appending it once filled would place it.
                richText = chartText.Append(chartText.New(C, "rich"));
            }
            if (richText.Child(A, "bodyPr").IsNone) {
                richText.Append(richText.New(A, "bodyPr"));
            }
            if (richText.Child(A, "lstStyle").IsNone) {
                richText.Append(richText.New(A, "lstStyle"));
            }

            OfficeXmlNode paragraph = richText.Child(A, "p");
            if (paragraph.IsNone) {
                paragraph = richText.Append(richText.New(A, "p"));
            }

            OfficeXmlNode run = paragraph.Child(A, "r");
            if (run.IsNone) {
                run = paragraph.Append(paragraph.New(A, "r"));
            }

            OfficeXmlNode runProps = run.Child(A, "rPr");
            if (runProps.IsNone) {
                runProps = run.InsertBefore(run.New(A, "rPr"), run.FirstChild);
            } else if (runProps != run.FirstChild) {
                runProps.Remove();
                runProps = run.InsertBefore(runProps, run.FirstChild);
            }

            return runProps;
        }

        private static OfficeXmlNode EnsureChartShapeProperties(OfficeXmlNode series) {
            OfficeXmlNode props = series.Child(C, "spPr");
            if (props.Exists) {
                props.Remove();
            } else {
                props = series.New(C, "spPr");
            }

            OfficeXmlNode anchor = series.Children(C, "tx").Last();
            if (anchor.IsNone) anchor = series.Children(C, "order").Last();
            if (anchor.IsNone) anchor = series.Children(C, "idx").Last();

            if (anchor.Exists) {
                return anchor.InsertAfterSelf(props);
            }
            return series.InsertBefore(props, series.FirstChild);
        }

        // Native theme slots remain linked; status roles are literal colors from the shared theme.
        private static OfficeXmlNode CreateChartSolidFill(OfficeXmlNode context, OfficeColor? color) {
            OfficeColor value = color ?? throw new ArgumentNullException(nameof(color));
            OfficeXmlNode fill = context.New(A, "solidFill");
            if (value.ThemeSlot is string slot) {
                fill.Append(fill.New(A, "schemeClr").Attr("val", slot));
                return fill;
            }
            value = value.Resolve(OfficeTheme.Default);
            OfficeXmlNode rgb = fill.New(A, "srgbClr").Attr("val", value.ToRgbHex());
            if (value.A != 255) rgb.Append(rgb.NewVal(A, "alpha", (int)Math.Round(value.A * 100000D / 255)));
            fill.Append(rgb);
            return fill;
        }

        private static void ApplySolidFill(OfficeXmlNode props, OfficeColor? color) {
            RemoveShapeFillChoices(props);
            InsertShapeFill(props, CreateChartSolidFill(props, color));
        }

        private static bool IsShapeFillChoice(string localName) =>
            localName == "noFill"
            || localName == "solidFill"
            || localName == "gradFill"
            || localName == "blipFill"
            || localName == "pattFill"
            || localName == "grpFill";

        private static bool FollowsShapeOutline(string localName) =>
            localName == "effectLst"
            || localName == "effectDag"
            || localName == "scene3d"
            || localName == "sp3d"
            || localName == "extLst";

        private static void RemoveShapeFillChoices(OfficeXmlNode props) {
            OfficeXmlNode child = props.FirstChild;
            while (child.Exists) {
                OfficeXmlNode next = child.NextSibling;
                if (IsShapeFillChoice(child.LocalName)) child.Remove();
                child = next;
            }
        }

        private static void InsertShapeFill(OfficeXmlNode props, OfficeXmlNode fill) {
            OfficeXmlNode insertBefore = OfficeXmlNode.None;
            foreach (OfficeXmlNode child in props.Children()) {
                if (child.Is(A, "ln") || FollowsShapeOutline(child.LocalName)) {
                    insertBefore = child;
                    break;
                }
            }
            props.InsertBefore(fill, insertBefore);
        }

        private static OfficeXmlNode InsertOutline(OfficeXmlNode props, OfficeXmlNode outline) {
            OfficeXmlNode insertBefore = OfficeXmlNode.None;
            foreach (OfficeXmlNode child in props.Children()) {
                if (FollowsShapeOutline(child.LocalName)) {
                    insertBefore = child;
                    break;
                }
            }
            return props.InsertBefore(outline, insertBefore);
        }

        private static void ApplyTextSolidFill(OfficeXmlNode props, OfficeColor? color) {
            props.RemoveChildren(A, "solidFill");
            props.RemoveChildren(A, "noFill");
            props.RemoveChildren(A, "gradFill");
            props.RemoveChildren(A, "pattFill");

            OfficeXmlNode fill = CreateChartSolidFill(props, color);
            OfficeXmlNode insertBefore = OfficeXmlNode.None;
            foreach (OfficeXmlNode child in props.Children()) {
                if (child.Is(A, "latin")
                    || child.Is(A, "ea")
                    || child.Is(A, "cs")
                    || child.Is(A, "sym")
                    || child.Is(A, "hlinkClick")
                    || child.Is(A, "hlinkMouseOver")
                    || child.Is(A, "extLst")) {
                    insertBefore = child;
                    break;
                }
            }

            props.InsertBefore(fill, insertBefore);
        }

        private static void ApplyLine(OfficeXmlNode props, OfficeColor? color, OfficeLength? width) {
            OfficeChartStyleBounds.ValidateLineWidth(width?.ToPoints(), nameof(width), allowZero: true);
            OfficeXmlNode existing = props.Child(A, "ln");
            OfficeXmlNode outline = existing.Exists ? existing : props.New(A, "ln");
            RemoveShapeFillChoices(outline);
            outline.InsertBefore(CreateChartSolidFill(outline, color), outline.FirstChild);
            if (width != null) {
                outline.SetAttribute("w", checked((int)width.Value.ToEmus()).ToString(CultureInfo.InvariantCulture));
            }

            if (existing.IsNone) {
                InsertOutline(props, outline);
            }
        }

        private static void ApplyOptionalLine(OfficeXmlNode props, OfficeColor? color, OfficeLength? width) {
            OfficeChartStyleBounds.ValidateLineWidth(width?.ToPoints(), nameof(width), allowZero: true);
            OfficeXmlNode existing = props.Child(A, "ln");
            OfficeXmlNode outline = existing.Exists ? existing : props.New(A, "ln");
            if (color != null) {
                RemoveShapeFillChoices(outline);
                outline.InsertBefore(CreateChartSolidFill(outline, color), outline.FirstChild);
            }
            if (width != null) {
                outline.SetAttribute("w", checked((int)width.Value.ToEmus()).ToString(CultureInfo.InvariantCulture));
            }
            if (existing.IsNone) {
                InsertOutline(props, outline);
            }
        }

        private static void ApplyMarker(OfficeXmlNode marker, OfficeChartMarkerShape? style, OfficeLength? size, OfficeColor? fillColor, OfficeColor? lineColor, OfficeLength? lineWidth) {
            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth), allowZero: true);
            if (size.HasValue && (size.Value < OfficeLength.Points(2) || size.Value > OfficeLength.Points(72))) throw new ArgumentOutOfRangeException(nameof(size), "Native marker sizes must be between 2 and 72.");
            if (style.HasValue) marker.SetOrdered(marker.NewVal(C, "symbol", style.Value.ToOpenXmlToken()), CMarkerView.Order);
            if (size != null) {
                byte nativeSize = checked((byte)decimal.Round(size.Value.ToEmus() / 12700m, 0, MidpointRounding.AwayFromZero));
                marker.SetOrdered(marker.NewVal(C, "size", (int)nativeSize), CMarkerView.Order);
            }

            if (fillColor != null || lineColor != null || lineWidth != null) {
                OfficeXmlNode existing = marker.Child(C, "spPr");
                OfficeXmlNode props = existing.Exists ? existing : marker.Append(marker.New(C, "spPr"));
                if (fillColor != null) {
                    ApplySolidFill(props, fillColor);
                }
                if (lineColor != null || lineWidth != null) {
                    OfficeColor? normalizedLine = lineColor != null ? lineColor : null;
                    ApplyOptionalLine(props, normalizedLine, lineWidth);
                }
            }
        }

    }
}
