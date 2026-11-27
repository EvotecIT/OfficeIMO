using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        /// <remarks>
        /// For a series not yet under a chart layer, <c>kind</c> is the chart kind whose layer it is built for; an
        /// attached series is read from its layer.
        /// </remarks>
        private static void SetAutomaticSeriesColor(OfficeXmlNode element, bool automatic, OfficeChartType? kind = null) {
            OfficeXmlNode extensions = default;
            foreach (OfficeXmlNode child in element.Children()) {
                if (child.LocalName == "extLst") {
                    extensions = child;
                    break;
                }
            }
            OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(element);
            if (!automatic) return;
            // Series of a surface layer, or elements that aren't chart series, carry no automatic color provenance.
            if (!element.Is(NsC, "ser") || SeriesLayer(element, kind) is not
                ("barChart" or "bar3DChart" or "lineChart" or "line3DChart" or "stockChart" or "scatterChart" or
                 "areaChart" or "area3DChart" or "radarChart" or "pieChart" or "pie3DChart" or "doughnutChart" or
                 "ofPieChart" or "bubbleChart")) return;
            if (extensions.IsNone) extensions = element.Append(element.CreateElement(NsC, "extLst"));
            OfficeXmlNode marker = extensions.Append(extensions.CreateElement(NsC, "ext"));
            marker.SetAttribute("uri", OfficeOpenXmlChartAutomaticColors.AutomaticSeriesColorExtensionUri);
            marker.Append(marker.CreateElement(OfficeOpenXmlChartAutomaticColors.AutomaticSeriesColorExtensionUri, "autoColor", "oi"));
        }

        private static OfficeColor? ReadDirectSeriesColor(OfficeXmlNode seriesElement) {
            OfficeXmlNode properties = seriesElement.Child(NsC, "spPr");
            OfficeXmlNode marker = seriesElement.Child(NsC, "marker").Child(NsC, "spPr");
            return OfficeOpenXmlThemeColorResolver.ResolveColor(properties.Child(NsA, "solidFill"), OfficeXmlNode.None)
                ?? OfficeOpenXmlThemeColorResolver.ResolveColor(properties.Child(NsA, "ln").Child(NsA, "solidFill"), OfficeXmlNode.None)
                ?? OfficeOpenXmlThemeColorResolver.ResolveColor(marker.Child(NsA, "solidFill"), OfficeXmlNode.None);
        }

        internal static void MaterializeDefaultSeriesColor(OfficeXmlNode element,
            OfficeChartType kind, int seriesIndex) {
            var appearance = new OfficeSeries(string.Empty, Array.Empty<double>(), null,
                OfficeChartStyle.Default.GetSeriesColor(seriesIndex));
            ApplySharedSeriesShapeStyle(element, appearance, kind, null);
            SetAutomaticSeriesColor(element, true, kind);
        }

        private static void ApplySharedSeriesShapeStyle(OfficeXmlNode seriesElement,
            OfficeSeries series, OfficeChartType kind,
            OfficeColor? fallbackFillColor, OfficeTheme? theme = null) {
            OfficeColor? fillColor = series.Color ?? fallbackFillColor;
            OfficeXmlNode properties = seriesElement.Child(NsC, "spPr");
            bool newProperties = properties.IsNone;
            if (newProperties) properties = seriesElement.CreateElement(NsC, "spPr");
            OfficeXmlNode existingOutline = properties.Child(NsA, "ln");
            OfficeColor? existingOutlineColor = OfficeOpenXmlThemeColorResolver.ResolveColor(
                existingOutline.Child(NsA, "solidFill"), OfficeXmlNode.None);
            OfficeColor? outlineColor = kind == OfficeChartType.Bubble
                ? series.MarkerOutlineColor ?? existingOutlineColor ?? fillColor
                : IsFilledSharedKind(kind) ? null : fillColor;
            OfficeLength? outlineWidth = kind == OfficeChartType.Bubble
                ? series.MarkerOutlineWidth ?? series.StrokeWidth
                : series.StrokeWidth;
            bool reenableBubbleOutline = kind == OfficeChartType.Bubble &&
                series.ShowMarkerOutline &&
                !existingOutline.Child(NsA, "noFill").IsNone;
            bool reenableConnectingLine = (!IsFilledSharedKind(kind) || IsAreaKind(kind)) && series.ConnectLine &&
                !existingOutline.Child(NsA, "noFill").IsNone;
            bool hideAreaOutline = IsAreaKind(kind) && !series.ConnectLine;
            if (!fillColor.HasValue && !outlineColor.HasValue && outlineWidth == null &&
                (kind != OfficeChartType.Bubble || series.ShowMarkerOutline) &&
                series.StrokeDashStyle == null &&
                (series.ConnectLine || IsFilledSharedKind(kind)) &&
                !reenableBubbleOutline && !reenableConnectingLine && !hideAreaOutline) return;
            if (fillColor.HasValue && (IsFilledSharedKind(kind) || kind == OfficeChartType.Line)) {
                RemoveSharedFillChoices(properties);
                SetOrdered(properties, CreateSharedSolidFill(properties, fillColor.Value, theme), DmlShapePropertiesView.Order);
            }
            OfficeXmlNode outline = properties.Child(NsA, "ln");
            bool newOutline = outline.IsNone;
            if (newOutline) outline = properties.CreateElement(NsA, "ln");
            bool replaceOutlineFill = reenableBubbleOutline || reenableConnectingLine ||
                (kind == OfficeChartType.Bubble && !series.ShowMarkerOutline) ||
                (!series.ConnectLine && !IsFilledSharedKind(kind)) ||
                hideAreaOutline ||
                outlineColor.HasValue;
            if (replaceOutlineFill) {
                RemoveSharedFillChoices(outline);
                if (kind == OfficeChartType.Bubble &&
                    !series.ShowMarkerOutline) {
                    SetOrdered(outline, outline.CreateElement(NsA, "noFill"), DmlLinePropertiesView.Order);
                } else if (hideAreaOutline || (!series.ConnectLine && !IsFilledSharedKind(kind))) {
                    SetOrdered(outline, outline.CreateElement(NsA, "noFill"), DmlLinePropertiesView.Order);
                } else if (outlineColor.HasValue) {
                    SetOrdered(outline, CreateSharedSolidFill(outline, outlineColor.Value, theme), DmlLinePropertiesView.Order);
                }
            }
            if (outlineWidth.HasValue) {
                outline.SetAttribute("w", Text(checked((int)outlineWidth.Value.ToEmus())));
            } else if (reenableConnectingLine && outline.IntAttr("w") == 0) {
                // A hidden native area outline may have zero width as well as no fill.
                outline.SetAttribute("w", null);
            }
            if (series.StrokeDashStyle.HasValue) {
                outline.RemoveChildren(NsA, "prstDash");
                SetOrdered(outline, outline.CreateVal(NsA, "prstDash", MapDash(series.StrokeDashStyle.Value)), DmlLinePropertiesView.Order);
            } else if (reenableConnectingLine || reenableBubbleOutline) {
                // A hidden imported outline can retain an unsupported dash. Re-enabling it
                // without an explicit shared dash means a solid native outline.
                outline.RemoveChildren(NsA, "prstDash");
            }
            // An empty a:ln reads back as an outline the caller never asked for, so a new
            // outline is only written when it carries a fill, width or dash.
            if (newOutline && (outline.HasChildElements() || outline.HasAttributes)) SetOrdered(properties, outline, DmlShapePropertiesView.Order);
            if (newProperties) InsertSharedSeriesProperties(seriesElement, properties, kind);
        }

        private static void ApplySharedSeriesMarker(OfficeXmlNode seriesElement,
            OfficeSeries series, OfficeChartType kind, OfficeColor? fallbackColor, OfficeTheme theme,
            bool preservePresentation) {
            if (!IsMarkerKind(kind)) return;
            OfficeXmlNode marker = seriesElement.Child(NsC, "marker");
            bool newMarker = marker.IsNone;
            if (newMarker) marker = seriesElement.CreateElement(NsC, "marker");
            if (!preservePresentation && !series.ResolveMarkerVisibility(kind, theme))
                SetOrdered(marker, marker.CreateVal(NsC, "symbol", "none"), CMarkerView.Order);
            else if (!preservePresentation && (series.ResolveMarkerShape(kind, theme).HasValue || marker.Child(NsC, "symbol").IsNone))
                SetOrdered(marker, marker.CreateVal(NsC, "symbol", MapMarker(series.ResolveMarkerShape(kind, theme))), CMarkerView.Order);
            if (series.MarkerSize.HasValue) {
                byte size = checked((byte)decimal.Round(OfficeChartStyleBounds.ClampNativeMarkerSize(series.MarkerSize.Value).ToEmus() / 12700m, 0, MidpointRounding.AwayFromZero));
                SetOrdered(marker, marker.CreateVal(NsC, "size", Text((int)size)), CMarkerView.Order);
            }
            OfficeColor? markerFill = series.Color ?? fallbackColor;
            if (markerFill.HasValue || series.MarkerOutlineColor.HasValue || series.MarkerOutlineWidth.HasValue) {
                OfficeXmlNode properties = marker.Child(NsC, "spPr");
                bool newProperties = properties.IsNone;
                if (newProperties) properties = marker.CreateElement(NsC, "spPr");
                if (markerFill.HasValue) {
                    RemoveSharedFillChoices(properties);
                    SetOrdered(properties, CreateSharedSolidFill(properties, markerFill.Value, theme), DmlShapePropertiesView.Order);
                }
                OfficeXmlNode outline = properties.Child(NsA, "ln");
                bool newOutline = outline.IsNone;
                if (newOutline) outline = properties.CreateElement(NsA, "ln");
                OfficeColor? markerColor = series.MarkerOutlineColor ?? markerFill;
                if (markerColor.HasValue) {
                    RemoveSharedFillChoices(outline);
                    SetOrdered(outline, CreateSharedSolidFill(outline, markerColor.Value, theme), DmlLinePropertiesView.Order);
                }
                if (series.MarkerOutlineWidth.HasValue) {
                    outline.SetAttribute("w", Text(checked((int)series.MarkerOutlineWidth.Value.ToEmus())));
                }
                if (newOutline) SetOrdered(properties, outline, DmlShapePropertiesView.Order);
                if (newProperties) SetOrdered(marker, properties, CMarkerView.Order);
            }
            if (newMarker) InsertSharedMarker(seriesElement, marker, kind);
        }

        private static void ApplySharedPointColors(OfficeXmlNode seriesElement, OfficeSeries series, OfficeTheme theme) {
            if (series.PointColors == null) return;
            OfficeXmlNode insertBefore = FirstExisting(seriesElement, "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal");
            var pointsByIndex = new Dictionary<uint, OfficeXmlNode>();
            foreach (OfficeXmlNode existingPoint in OfficeOpenXmlChartPointStyles.GetBoundedPoints(seriesElement)) {
                uint? existingIndex = existingPoint.Child(NsC, "idx").UIntVal();
                if (existingIndex.HasValue && !pointsByIndex.ContainsKey(existingIndex.Value)) {
                    pointsByIndex.Add(existingIndex.Value, existingPoint);
                }
            }
            for (int index = 0; index < series.PointColors.Count; index++) {
                OfficeColor? color = series.PointColors[index];
                uint pointIndex = (uint)index;
                if (!pointsByIndex.TryGetValue(pointIndex, out OfficeXmlNode point)) {
                    if (!color.HasValue) continue;
                    point = seriesElement.CreateElement(NsC, "dPt");
                    point.AppendVal(NsC, "idx", Text((uint)index));
                    point = seriesElement.InsertBefore(point, insertBefore);
                    pointsByIndex.Add(pointIndex, point);
                }
                OfficeXmlNode direct = point.Child(NsC, "spPr");
                OfficeXmlNode marker = point.Child(NsC, "marker").Child(NsC, "spPr");
                if (!direct.IsNone) RemoveSharedFillChoices(direct);
                if (!marker.IsNone) RemoveSharedFillChoices(marker);
                if (color.HasValue) {
                    OfficeXmlNode properties = !marker.IsNone ? marker : direct;
                    bool newProperties = properties.IsNone;
                    if (newProperties) properties = point.CreateElement(NsC, "spPr");
                    SetOrdered(properties, CreateSharedSolidFill(properties, color.Value, theme), DmlShapePropertiesView.Order);
                    if (newProperties) SetOrdered(point, properties, CDPtView.Order);
                }
                if (!direct.IsNone && !direct.HasChildElements() && !direct.HasAttributes) direct.Remove();
                if (!marker.IsNone && !marker.HasChildElements() && !marker.HasAttributes) marker.Remove();
            }
        }

        // The first of these children that exists, by priority rather than document order.
        private static OfficeXmlNode FirstExisting(OfficeXmlNode parent, params string[] names) {
            foreach (string name in names) {
                OfficeXmlNode child = parent.Child(NsC, name);
                if (!child.IsNone) return child;
            }
            return default;
        }

        private static void RemoveSharedFillChoices(OfficeXmlNode element) {
            element.RemoveChildren(NsA, "noFill");
            element.RemoveChildren(NsA, "solidFill");
            element.RemoveChildren(NsA, "gradFill");
            element.RemoveChildren(NsA, "pattFill");
            element.RemoveChildren(NsA, "blipFill");
            element.RemoveChildren(NsA, "grpFill");
        }

        /// <summary>A detached <c>a:solidFill</c> holding <see cref="CreateSharedRgbColor"/>, built for <paramref name="context"/>.</summary>
        private static OfficeXmlNode CreateSharedSolidFill(OfficeXmlNode context, OfficeColor color, OfficeTheme? theme = null) {
            OfficeXmlNode fill = context.CreateElement(NsA, "solidFill");
            fill.Append(CreateSharedRgbColor(fill, color, theme));
            return fill;
        }

        private static OfficeXmlNode CreateSharedRgbColor(OfficeXmlNode context, OfficeColor color, OfficeTheme? theme = null) {
            color = color.Resolve(theme ?? OfficeTheme.Default);
            OfficeXmlNode rgb = context.CreateVal(NsA, "srgbClr", color.ToRgbHex());
            if (color.A < byte.MaxValue) {
                rgb.AppendVal(NsA, "alpha", Text(checked((int)Math.Round(
                    color.A / 255D * 100000D,
                    MidpointRounding.AwayFromZero))));
            }
            return rgb;
        }

        // The series kind stands in for the SDK series class when the series isn't under a layer yet.
        private static void InsertSharedSeriesProperties(OfficeXmlNode series, OfficeXmlNode properties, OfficeChartType kind) =>
            SetSeriesChild(series, properties, kind);

        private static void InsertSharedMarker(OfficeXmlNode series, OfficeXmlNode marker, OfficeChartType kind) =>
            SetSeriesChild(series, marker, kind);

        private static void SetSeriesChild(OfficeXmlNode series, OfficeXmlNode child, OfficeChartType? kind) {
            OfficeXmlChildOrder? order = SeriesOrder(SeriesLayer(series, kind));
            if (order != null) SetOrdered(series, child, order);
            else series.Append(child);
        }

        // SDK callers

        internal static void MaterializeDefaultSeriesColor(OpenXmlCompositeElement element,
            OfficeChartType kind, int seriesIndex) =>
            MaterializeDefaultSeriesColor(OfficeXmlNode.FromSdk(element), kind, seriesIndex);
    }
}
