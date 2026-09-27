using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        private static void ApplySharedSeriesShapeStyle(OpenXmlCompositeElement seriesElement,
            OfficeChartSeries series, OfficeChartKind kind,
            OfficeColor? fallbackFillColor) {
            OfficeColor? fillColor = series.Color ?? fallbackFillColor;
            OfficeColor? outlineColor = kind == OfficeChartKind.Bubble
                ? series.MarkerOutlineColor ?? fillColor
                : fillColor;
            double? outlineWidth = kind == OfficeChartKind.Bubble
                ? series.MarkerOutlineWidth ?? series.StrokeWidth
                : series.StrokeWidth;
            C.ChartShapeProperties properties =
                seriesElement.GetFirstChild<C.ChartShapeProperties>() ??
                new C.ChartShapeProperties();
            bool reenableBubbleOutline = kind == OfficeChartKind.Bubble &&
                series.ShowMarkerOutline &&
                properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.NoFill>() != null;
            bool reenableConnectingLine = !IsFilledSharedKind(kind) && series.ConnectLine &&
                properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.NoFill>() != null;
            if (!fillColor.HasValue && !outlineColor.HasValue && outlineWidth == null &&
                (kind != OfficeChartKind.Bubble || series.ShowMarkerOutline) &&
                series.StrokeDashStyle == null &&
                (series.ConnectLine || IsFilledSharedKind(kind)) &&
                !reenableBubbleOutline && !reenableConnectingLine) return;
            if (fillColor.HasValue && IsFilledSharedKind(kind)) {
                properties.RemoveAllChildren<A.SolidFill>();
                properties.RemoveAllChildren<A.NoFill>();
                properties.RemoveAllChildren<A.GradientFill>();
                properties.RemoveAllChildren<A.PatternFill>();
                properties.RemoveAllChildren<A.BlipFill>();
                properties.RemoveAllChildren<A.GroupFill>();
                properties.AddChild(new A.SolidFill(CreateSharedRgbColor(fillColor.Value)), true);
            }
            A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
            bool replaceOutlineFill = reenableBubbleOutline || reenableConnectingLine ||
                (kind == OfficeChartKind.Bubble && !series.ShowMarkerOutline) ||
                (!series.ConnectLine && !IsFilledSharedKind(kind)) ||
                outlineColor.HasValue;
            if (replaceOutlineFill) {
                outline.RemoveAllChildren<A.SolidFill>();
                outline.RemoveAllChildren<A.NoFill>();
                outline.RemoveAllChildren<A.GradientFill>();
                outline.RemoveAllChildren<A.PatternFill>();
                outline.RemoveAllChildren<A.BlipFill>();
                outline.RemoveAllChildren<A.GroupFill>();
                if (kind == OfficeChartKind.Bubble &&
                    !series.ShowMarkerOutline) {
                    outline.AddChild(new A.NoFill(), true);
                } else if (!series.ConnectLine &&
                           !IsFilledSharedKind(kind)) {
                    outline.AddChild(new A.NoFill(), true);
                } else if (outlineColor.HasValue) {
                    outline.AddChild(new A.SolidFill(CreateSharedRgbColor(outlineColor.Value)), true);
                }
            }
            if (outlineWidth.HasValue) {
                outline.Width = checked((int)FromPoints(outlineWidth.Value));
            }
            if (series.StrokeDashStyle.HasValue) {
                outline.RemoveAllChildren<A.PresetDash>();
                outline.AddChild(new A.PresetDash { Val = MapDash(series.StrokeDashStyle.Value) }, true);
            }
            if (outline.Parent == null) properties.AddChild(outline, true);
            if (properties.Parent == null) InsertSharedSeriesProperties(seriesElement, properties);
        }

        private static void ApplySharedSeriesMarker(OpenXmlCompositeElement seriesElement,
            OfficeChartSeries series, OfficeChartKind kind) {
            if (!IsMarkerKind(kind)) return;
            C.Marker marker = seriesElement.GetFirstChild<C.Marker>() ?? new C.Marker();
            marker.Symbol = new C.Symbol {
                Val = series.ShowMarkers ? MapMarker(series.MarkerShape) : C.MarkerStyleValues.None
            };
            if (series.MarkerSize.HasValue) marker.Size = new C.Size { Val = (byte)OfficeChartStyleBounds.ClampNativeMarkerSize(series.MarkerSize.Value) };
            if (series.Color.HasValue || series.MarkerOutlineColor.HasValue || series.MarkerOutlineWidth.HasValue) {
                C.ChartShapeProperties properties = marker.ChartShapeProperties ?? new C.ChartShapeProperties();
                if (series.Color.HasValue) {
                    RemoveSharedFillChoices(properties);
                    properties.AddChild(new A.SolidFill(CreateSharedRgbColor(series.Color.Value)), true);
                }
                A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
                OfficeColor? markerColor = series.MarkerOutlineColor ?? series.Color;
                if (markerColor.HasValue) {
                    RemoveSharedFillChoices(outline);
                    outline.AddChild(new A.SolidFill(CreateSharedRgbColor(markerColor.Value)), true);
                }
                if (series.MarkerOutlineWidth.HasValue) {
                    outline.Width = checked((int)FromPoints(series.MarkerOutlineWidth.Value));
                }
                if (outline.Parent == null) properties.AddChild(outline, true);
                if (properties.Parent == null) marker.AddChild(properties, true);
            }
            if (marker.Parent == null) InsertSharedMarker(seriesElement, marker);
        }

        private static void ApplySharedPointColors(OpenXmlCompositeElement seriesElement, OfficeChartSeries series) {
            if (series.PointColors == null) return;
            OpenXmlElement? insertBefore = seriesElement.GetFirstChild<C.DataLabels>() ??
                (OpenXmlElement?)seriesElement.GetFirstChild<C.Trendline>() ??
                (OpenXmlElement?)seriesElement.GetFirstChild<C.ErrorBars>() ??
                (OpenXmlElement?)seriesElement.GetFirstChild<C.CategoryAxisData>() ??
                (OpenXmlElement?)seriesElement.GetFirstChild<C.Values>() ??
                (OpenXmlElement?)seriesElement.GetFirstChild<C.XValues>() ?? seriesElement.GetFirstChild<C.YValues>();
            var pointsByIndex = new Dictionary<uint, C.DataPoint>();
            foreach (C.DataPoint existingPoint in seriesElement.Elements<C.DataPoint>()) {
                uint? existingIndex = existingPoint.Index?.Val?.Value;
                if (existingIndex.HasValue && !pointsByIndex.ContainsKey(existingIndex.Value)) {
                    pointsByIndex.Add(existingIndex.Value, existingPoint);
                }
            }
            for (int index = 0; index < series.PointColors.Count; index++) {
                OfficeColor? color = series.PointColors[index];
                if (!color.HasValue) continue;
                uint pointIndex = (uint)index;
                if (!pointsByIndex.TryGetValue(pointIndex, out C.DataPoint? point)) {
                    point = new C.DataPoint(new C.Index { Val = (uint)index });
                    if (insertBefore != null) {
                        seriesElement.InsertBefore(point, insertBefore);
                    } else {
                        seriesElement.Append(point);
                    }
                    pointsByIndex.Add(pointIndex, point);
                }
                C.ChartShapeProperties? properties =
                    point.GetFirstChild<C.ChartShapeProperties>();
                if (properties == null) {
                    properties = new C.ChartShapeProperties();
                    point.AddChild(properties, true);
                }
                properties.RemoveAllChildren<A.SolidFill>();
                properties.RemoveAllChildren<A.NoFill>();
                properties.RemoveAllChildren<A.GradientFill>();
                properties.RemoveAllChildren<A.PatternFill>();
                properties.RemoveAllChildren<A.BlipFill>();
                properties.RemoveAllChildren<A.GroupFill>();
                properties.AddChild(new A.SolidFill(CreateSharedRgbColor(color.Value)), true);
            }
        }

        private static void RemoveSharedFillChoices(OpenXmlCompositeElement element) {
            element.RemoveAllChildren<A.NoFill>();
            element.RemoveAllChildren<A.SolidFill>();
            element.RemoveAllChildren<A.GradientFill>();
            element.RemoveAllChildren<A.PatternFill>();
            element.RemoveAllChildren<A.BlipFill>();
            element.RemoveAllChildren<A.GroupFill>();
        }

        private static A.RgbColorModelHex CreateSharedRgbColor(OfficeColor color) {
            var rgb = new A.RgbColorModelHex { Val = color.ToRgbHex() };
            if (color.A < byte.MaxValue) {
                rgb.Append(new A.Alpha {
                    Val = checked((int)Math.Round(
                        color.A / 255D * 100000D,
                        MidpointRounding.AwayFromZero))
                });
            }
            return rgb;
        }

        private static void InsertSharedSeriesProperties(OpenXmlCompositeElement series, C.ChartShapeProperties properties) {
            OpenXmlElement? insertBefore = series.GetFirstChild<C.InvertIfNegative>() ??
                (OpenXmlElement?)series.GetFirstChild<C.Marker>() ?? series.GetFirstChild<C.DataPoint>() ??
                series.GetFirstChild<C.DataLabels>() ?? series.GetFirstChild<C.CategoryAxisData>() ??
                (OpenXmlElement?)series.GetFirstChild<C.Values>() ??
                (OpenXmlElement?)series.GetFirstChild<C.XValues>() ?? series.GetFirstChild<C.YValues>();
            if (insertBefore != null) series.InsertBefore(properties, insertBefore);
            else series.Append(properties);
        }

        private static void InsertSharedMarker(OpenXmlCompositeElement series, C.Marker marker) {
            OpenXmlElement? insertBefore = series.GetFirstChild<C.DataPoint>() ??
                (OpenXmlElement?)series.GetFirstChild<C.DataLabels>() ?? series.GetFirstChild<C.CategoryAxisData>() ??
                (OpenXmlElement?)series.GetFirstChild<C.Values>() ??
                (OpenXmlElement?)series.GetFirstChild<C.XValues>() ?? series.GetFirstChild<C.YValues>();
            if (insertBefore != null) series.InsertBefore(marker, insertBefore);
            else series.Append(marker);
        }

    }
}
