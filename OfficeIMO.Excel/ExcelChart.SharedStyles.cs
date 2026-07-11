using System.Collections.Generic;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        internal void ApplyAuthoredSeriesStyles(IReadOnlyList<ExcelChartSeries> seriesStyles,
            IReadOnlyList<bool>? seriesLegendVisibility = null) {
            bool changed = false;
            for (int seriesIndex = 0; seriesIndex < seriesStyles.Count; seriesIndex++) {
                ExcelChartSeries style = seriesStyles[seriesIndex];
                changed |= ApplySeriesByIndex(seriesIndex, series => ApplyAuthoredSeriesStyle(series, style));

                if (style.PointColorArgb != null) {
                    for (int pointIndex = 0; pointIndex < style.PointColorArgb.Count; pointIndex++) {
                        string? color = style.PointColorArgb[pointIndex];
                        if (!string.IsNullOrWhiteSpace(color)) {
                            int currentPoint = pointIndex;
                            string currentColor = color!;
                            changed |= ApplySeriesByIndex(seriesIndex,
                                series => ApplyPointFill(series, currentPoint, NormalizeHexColor(currentColor)));
                        }
                    }
                }

                changed |= ApplySeriesMarkerByIndex(seriesIndex, marker => ApplyMarker(
                    marker,
                    style.ShowMarkers ? MapMarkerStyle(style.MarkerShape) : C.MarkerStyleValues.None,
                    style.MarkerSize,
                    style.ShowMarkers ? style.SeriesColorArgb : null,
                    style.ShowMarkers ? style.MarkerOutlineColorArgb : null,
                    style.ShowMarkers ? style.MarkerOutlineWidth : null));
            }

            if (seriesLegendVisibility != null) {
                changed |= ApplySeriesLegendVisibility(seriesLegendVisibility);
            }

            if (changed) Save();
        }

        private bool ApplySeriesLegendVisibility(IReadOnlyList<bool> seriesLegendVisibility) {
            C.Chart chart = GetChart();
            C.Legend? legend = chart.GetFirstChild<C.Legend>();
            bool hasHiddenSeries = false;
            for (int index = 0; index < seriesLegendVisibility.Count; index++) {
                if (!seriesLegendVisibility[index]) {
                    hasHiddenSeries = true;
                    break;
                }
            }
            if (legend == null && !hasHiddenSeries) return false;

            if (legend == null) {
                legend = new C.Legend(
                    new C.LegendPosition { Val = C.LegendPositionValues.Bottom },
                    new C.Layout(),
                    new C.Overlay { Val = false });
                C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
                if (plotArea != null) chart.InsertAfter(legend, plotArea);
                else chart.Append(legend);
            }

            bool changed = false;
            C.LegendEntry? existing;
            while ((existing = legend.GetFirstChild<C.LegendEntry>()) != null) {
                existing.Remove();
                changed = true;
            }
            for (int index = 0; index < seriesLegendVisibility.Count; index++) {
                if (seriesLegendVisibility[index]) continue;
                var entry = new C.LegendEntry(new C.Index { Val = (uint)index }, new C.Delete { Val = true });
                C.LegendPosition? position = legend.GetFirstChild<C.LegendPosition>();
                if (position != null) legend.InsertBefore(entry, position);
                else legend.PrependChild(entry);
                changed = true;
            }
            return changed;
        }

        private static void ApplyAuthoredSeriesStyle(OpenXmlCompositeElement series,
            ExcelChartSeries style) {
            bool hasShapeStyle = !string.IsNullOrWhiteSpace(style.SeriesColorArgb) ||
                                 style.SeriesLineWidth.HasValue || style.SeriesLineDashStyle.HasValue ||
                                 !style.ConnectLine;
            if (!hasShapeStyle) return;

            C.ChartShapeProperties properties = EnsureChartShapeProperties(series);
            if (!string.IsNullOrWhiteSpace(style.SeriesColorArgb)) {
                string color = NormalizeHexColor(style.SeriesColorArgb!);
                ApplySolidFill(properties, color);
                ApplyOptionalLine(properties, color, style.SeriesLineWidth);
            } else if (style.SeriesLineWidth.HasValue) {
                ApplyOptionalLine(properties, null, style.SeriesLineWidth);
            }

            if (style.SeriesLineDashStyle.HasValue) {
                A.Outline outline = properties.GetFirstChild<A.Outline>() ?? properties.AppendChild(new A.Outline());
                outline.RemoveAllChildren<A.PresetDash>();
                outline.Append(new A.PresetDash { Val = MapDashStyle(style.SeriesLineDashStyle.Value) });
            }
            if (!style.ConnectLine) ApplyNoLine(properties);
        }

        private static C.MarkerStyleValues MapMarkerStyle(OfficeChartMarkerShape? shape) {
            switch (shape ?? OfficeChartMarkerShape.Circle) {
                case OfficeChartMarkerShape.Square: return C.MarkerStyleValues.Square;
                case OfficeChartMarkerShape.Diamond: return C.MarkerStyleValues.Diamond;
                case OfficeChartMarkerShape.Triangle: return C.MarkerStyleValues.Triangle;
                case OfficeChartMarkerShape.Dash: return C.MarkerStyleValues.Dash;
                case OfficeChartMarkerShape.Dot: return C.MarkerStyleValues.Dot;
                case OfficeChartMarkerShape.Plus: return C.MarkerStyleValues.Plus;
                case OfficeChartMarkerShape.X: return C.MarkerStyleValues.X;
                case OfficeChartMarkerShape.Star: return C.MarkerStyleValues.Star;
                default: return C.MarkerStyleValues.Circle;
            }
        }

        private static A.PresetLineDashValues MapDashStyle(OfficeStrokeDashStyle style) {
            switch (style) {
                case OfficeStrokeDashStyle.Dash: return A.PresetLineDashValues.Dash;
                case OfficeStrokeDashStyle.Dot: return A.PresetLineDashValues.Dot;
                case OfficeStrokeDashStyle.DashDot: return A.PresetLineDashValues.DashDot;
                case OfficeStrokeDashStyle.DashDotDot: return A.PresetLineDashValues.LargeDashDotDot;
                default: return A.PresetLineDashValues.Solid;
            }
        }
    }
}
