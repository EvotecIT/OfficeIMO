using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        internal void ApplyAuthoredSeriesStyles(IReadOnlyList<ExcelChartSeries> seriesStyles,
            IReadOnlyList<bool>? seriesLegendVisibility = null, IReadOnlyList<OfficeSeries>? authoredSeries = null) {
            bool changed = false;
            for (int seriesIndex = 0; seriesIndex < seriesStyles.Count; seriesIndex++) {
                ExcelChartSeries style = seriesStyles[seriesIndex];
                OfficeSeries? authored = authoredSeries?[seriesIndex];
                changed |= ApplySeriesByChartIndex(seriesIndex,
                    (OfficeXmlNode series) => ApplyAuthoredSeriesStyle(series, style, authored?.Color));

                if (style.PointColorArgb != null) {
                    for (int pointIndex = 0; pointIndex < style.PointColorArgb.Count; pointIndex++) {
                        string? color = style.PointColorArgb[pointIndex];
                        if (!string.IsNullOrWhiteSpace(color)) {
                            int currentPoint = pointIndex;
                            string currentColor = color!;
                            changed |= ApplySeriesByChartIndex(seriesIndex,
                                (OfficeXmlNode series) => ApplyPointFill(series, currentPoint, GetAuthoredColor(authored?.PointColors?[currentPoint], currentColor)));
                        }
                    }
                }

                changed |= ApplySeriesMarkerByChartIndex(seriesIndex, (OfficeXmlNode marker) => ApplyMarker(
                    marker,
                    style.ShowMarkers ? MapMarkerStyle(style.MarkerShape) : OfficeChartMarkerShape.None,
                    style.MarkerSize.HasValue ? OfficeChartStyleBounds.ClampNativeMarkerSize(style.MarkerSize.Value) : (OfficeLength?)null,
                    style.ShowMarkers && style.SeriesColorArgb != null ? GetAuthoredColor(authored?.Color, style.SeriesColorArgb) : (OfficeColor?)null,
                    style.ShowMarkers && style.MarkerOutlineColorArgb != null ? GetAuthoredColor(authored?.MarkerOutlineColor, style.MarkerOutlineColorArgb) : (OfficeColor?)null,
                    style.ShowMarkers ? style.MarkerOutlineWidth : null));
            }

            if (seriesLegendVisibility != null) {
                changed |= ApplySeriesLegendVisibility(seriesLegendVisibility, seriesStyles);
            }

            if (changed) Save();
        }

        private bool ApplySeriesLegendVisibility(IReadOnlyList<bool> seriesLegendVisibility, IReadOnlyList<ExcelChartSeries> seriesStyles) {
            OfficeXmlNode chart = ChartNode();
            OfficeXmlNode legend = chart.Child(C, "legend");
            bool hasHiddenSeries = false;
            for (int index = 0; index < seriesLegendVisibility.Count; index++) {
                if (!seriesLegendVisibility[index]) {
                    hasHiddenSeries = true;
                    break;
                }
            }
            if (legend.IsNone && !hasHiddenSeries) return false;

            if (legend.IsNone) {
                legend = chart.New(C, "legend");
                legend.With(
                    legend.NewVal(C, "legendPos", "b"),
                    legend.New(C, "layout"),
                    legend.NewVal(C, "overlay", false));
                OfficeXmlNode plotArea = chart.Child(C, "plotArea");
                legend = plotArea.Exists ? plotArea.InsertAfterSelf(legend) : chart.Append(legend);
            }

            bool changed = false;
            OfficeXmlNode existing;
            while ((existing = legend.Child(C, "legendEntry")).Exists) {
                existing.Remove();
                changed = true;
            }
            var trailing = new List<OfficeXmlNode>();
            for (OfficeXmlNode child = legend.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (!child.Is(C, "legendPos")) trailing.Add(child);
            }
            foreach (OfficeXmlNode child in trailing) child.Remove();
            bool radial = ChartType == OfficeChartType.Pie || ChartType == OfficeChartType.Doughnut;
            int entryCount = radial ? (seriesStyles.Count == 0 ? 0 : seriesStyles[0].Values.Count) : seriesLegendVisibility.Count;
            for (int index = 0; index < entryCount; index++) {
                if (radial ? !hasHiddenSeries : seriesLegendVisibility[index]) continue;
                OfficeXmlNode entry = legend.New(C, "legendEntry");
                entry.With(entry.NewVal(C, "idx", (uint)index), entry.NewVal(C, "delete", true));
                legend.Append(entry);
                changed = true;
            }
            foreach (OfficeXmlNode child in trailing) legend.Append(child);
            return changed;
        }

        private static void ApplyAuthoredSeriesStyle(OfficeXmlNode series,
            ExcelChartSeries style, OfficeColor? authoredColor = null) {
            bool reenableLine = style.ConnectLine && IsLineScatterOrRadarSeries(series) &&
                series.Child(C, "spPr").Child(A, "ln").Child(A, "noFill").Exists;
            bool hasShapeStyle = !string.IsNullOrWhiteSpace(style.SeriesColorArgb) ||
                                 style.SeriesLineWidth.HasValue || style.SeriesLineDashStyle.HasValue ||
                                 !style.ConnectLine || reenableLine;
            if (!hasShapeStyle && style.SeriesStyle == null) return;

            OfficeXmlNode properties = EnsureChartShapeProperties(series);
            if (reenableLine) properties.Child(A, "ln").RemoveChildren(A, "noFill");
            if (!string.IsNullOrWhiteSpace(style.SeriesColorArgb)) {
                OfficeColor color = GetAuthoredColor(authoredColor, style.SeriesColorArgb!);
                ApplySolidFill(properties, color);
                ApplyOptionalLine(properties, color, style.SeriesLineWidth);
            } else if (style.SeriesLineWidth.HasValue) {
                ApplyOptionalLine(properties, null, style.SeriesLineWidth);
            }

            if (style.SeriesLineDashStyle.HasValue) {
                OfficeXmlNode outline = properties.Child(A, "ln");
                if (outline.IsNone) outline = properties.Append(properties.New(A, "ln"));
                outline.RemoveChildren(A, "prstDash");
                outline.Append(outline.NewVal(A, "prstDash", MapDashStyle(style.SeriesLineDashStyle.Value)));
            }
            if (!style.ConnectLine) ApplyNoLine(properties);
            if (style.SeriesStyle != null) OfficeOpenXmlChartPointStyles.ApplySeriesPaint(properties, style.SeriesStyle);
        }

        /// <summary>Whether a series is a line (including stock and 3-D line), scatter or radar <c>c:ser</c>.</summary>
        private static bool IsLineScatterOrRadarSeries(OfficeXmlNode series) {
            if (!series.Is(C, "ser")) return false;
            OfficeXmlNode parent = series.Parent;
            return parent.Is(C, "lineChart") || parent.Is(C, "line3DChart") || parent.Is(C, "stockChart")
                || parent.Is(C, "scatterChart") || parent.Is(C, "radarChart");
        }

        // Slots stay native; metadata already contains creation-theme-resolved status roles and literal ARGB.
        private static OfficeColor GetAuthoredColor(OfficeColor? authored, string argb) =>
            authored?.ThemeSlot != null ? authored.Value
                : ColorFromArgb(argb)!.Value;

        private static OfficeChartMarkerShape MapMarkerStyle(OfficeChartMarkerShape? shape) {
            switch (shape ?? OfficeChartMarkerShape.Circle) {
                case OfficeChartMarkerShape.Square: return OfficeChartMarkerShape.Square;
                case OfficeChartMarkerShape.Diamond: return OfficeChartMarkerShape.Diamond;
                case OfficeChartMarkerShape.Triangle: return OfficeChartMarkerShape.Triangle;
                case OfficeChartMarkerShape.Dash: return OfficeChartMarkerShape.Dash;
                case OfficeChartMarkerShape.Dot: return OfficeChartMarkerShape.Dot;
                case OfficeChartMarkerShape.Plus: return OfficeChartMarkerShape.Plus;
                case OfficeChartMarkerShape.X: return OfficeChartMarkerShape.X;
                case OfficeChartMarkerShape.Star: return OfficeChartMarkerShape.Star;
                default: return OfficeChartMarkerShape.Circle;
            }
        }

        /// <summary>The <c>a:prstDash</c> token for a stroke dash style.</summary>
        private static string MapDashStyle(OfficeStrokeDashStyle style) {
            switch (style) {
                case OfficeStrokeDashStyle.Dash: return "dash";
                case OfficeStrokeDashStyle.Dot: return "dot";
                case OfficeStrokeDashStyle.DashDot: return "dashDot";
                case OfficeStrokeDashStyle.DashDotDot: return "lgDashDotDot";
                default: return "solid";
            }
        }
    }
}
