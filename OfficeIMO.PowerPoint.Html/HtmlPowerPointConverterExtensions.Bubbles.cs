using System;
using System.Collections.Generic;
using System.Globalization;
using AngleSharp.Dom;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    private static bool TryCreateBubbleChartData(
        PptCore.PowerPointChartData data,
        out OfficeChartData? bubbleData) {
        bubbleData = null;
        var series = new List<OfficeChartSeries>();
        foreach (PptCore.PowerPointChartSeries item in data.Series) {
            if (item.XValues == null || item.BubbleSizes == null ||
                item.XValues.Count != item.Values.Count ||
                item.BubbleSizes.Count != item.Values.Count) {
                return false;
            }

            series.Add(OfficeChartSeries.CreateBubble(
                item.Name, item.XValues, item.Values, item.BubbleSizes,
                item.Color, item.PointColors,
                showInLegend: item.ShowInLegend,
                markerOutlineColor: item.StrokeColor ?? item.Color,
                markerOutlineWidth: item.StrokeWidth,
                showMarkerOutline: item.ShowStroke));
        }

        if (series.Count == 0) {
            return false;
        }

        bubbleData = new OfficeChartData(data.Categories, series);
        return true;
    }

    private static bool TryReadBubbleSizing(
        IElement item, out uint scalePercent,
        out OfficeChartBubbleSizeMode sizeMode) {
        scalePercent = 100U;
        sizeMode = OfficeChartBubbleSizeMode.Area;
        IElement? table =
            item.QuerySelector("table.officeimo-chart-data");
        if (table == null) {
            return true;
        }

        string? rawScale =
            table.GetAttribute("data-officeimo-bubble-scale");
        if (rawScale != null &&
            (!uint.TryParse(
                 rawScale, NumberStyles.None,
                 CultureInfo.InvariantCulture, out scalePercent) ||
             scalePercent > 300U)) {
            return false;
        }

        string? rawMode =
            table.GetAttribute("data-officeimo-bubble-size-mode");
        return rawMode == null ||
               Enum.TryParse(
                   rawMode, ignoreCase: true, out sizeMode) &&
               Enum.IsDefined(
                   typeof(OfficeChartBubbleSizeMode), sizeMode);
    }

    private static bool TryReadBubbleLegend(
        IElement item, out bool showLegend,
        out OfficeChartLegendPosition position,
        out bool overlayLegend) {
        showLegend = true;
        position = OfficeChartLegendPosition.Bottom;
        overlayLegend = false;
        IElement? table =
            item.QuerySelector("table.officeimo-chart-data");
        if (table == null) {
            return true;
        }

        string? rawShow =
            table.GetAttribute("data-officeimo-show-legend");
        if (rawShow != null &&
            !bool.TryParse(rawShow, out showLegend)) {
            return false;
        }

        string? rawPosition =
            table.GetAttribute("data-officeimo-legend-position");
        if (rawPosition != null &&
            (!Enum.TryParse(
                 rawPosition, ignoreCase: true, out position) ||
             !Enum.IsDefined(
                 typeof(OfficeChartLegendPosition), position))) {
            return false;
        }

        string? rawOverlay =
            table.GetAttribute("data-officeimo-overlay-legend");
        return rawOverlay == null ||
               bool.TryParse(rawOverlay, out overlayLegend);
    }

    private static PptCore.PowerPointChartLegendPosition ToPowerPointLegendPosition(
        OfficeChartLegendPosition position) =>
        position switch {
            OfficeChartLegendPosition.Left =>
                PptCore.PowerPointChartLegendPosition.Left,
            OfficeChartLegendPosition.Right =>
                PptCore.PowerPointChartLegendPosition.Right,
            OfficeChartLegendPosition.Top =>
                PptCore.PowerPointChartLegendPosition.Top,
            OfficeChartLegendPosition.Bottom =>
                PptCore.PowerPointChartLegendPosition.Bottom,
            _ => throw new ArgumentOutOfRangeException(nameof(position))
        };
}
