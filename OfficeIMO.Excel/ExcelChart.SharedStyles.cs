using System.Collections.Generic;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        internal void ApplyAuthoredSeriesStyles(IReadOnlyList<ExcelChartSeries> seriesStyles) {
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

            if (changed) Save();
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
