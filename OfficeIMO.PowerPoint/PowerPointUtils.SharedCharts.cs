using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
        internal const int MaximumSharedChartPoints = OfficeOpenXmlChartWriter.MaximumSharedChartPoints;
        internal static void ValidateSharedChartData(OfficeChartData data, OfficeChartKind kind) =>
            OfficeOpenXmlChartWriter.ValidateSharedChartData(data, kind);
        internal static void ValidateSharedWorkbookDimensions(int categoryCount, int seriesCount, long totalPoints) =>
            OfficeOpenXmlChartWriter.ValidateSharedWorkbookDimensions(categoryCount, seriesCount, totalPoints);
        internal static void ValidateBubbleWorkbookDimensions(int seriesCount, int maximumPoints, long totalPoints) =>
            OfficeOpenXmlChartWriter.ValidateBubbleWorkbookDimensions(seriesCount, maximumPoints, totalPoints);
        internal static void PopulateSharedChart(ChartPart part, string relationship, OfficeChartData data, OfficeChartKind kind) =>
            OfficeOpenXmlChartWriter.PopulateSharedChart(part, relationship, data, kind);
        internal static void UpdateSharedChartData(ChartPart part, OfficeChartData data, OfficeChartKind kind) =>
            OfficeOpenXmlChartWriter.UpdateSharedChartData(part, data, kind);
        internal static byte[] BuildBubbleChartWorkbook(OfficeChartData data) =>
            OfficeOpenXmlChartWriter.BuildWorkbook(data, OfficeChartKind.Bubble);
        internal static PowerPointChartData ToPowerPointChartData(OfficeChartData data) =>
            new(data.Categories, data.Series.Select(series =>
                new PowerPointChartSeries(series.Name, series.Values)));

        internal static PowerPointScatterChartData ToPowerPointScatterChartData(OfficeChartData data) {
            IReadOnlyList<double>? sharedX = data.Series.Any(series => series.XValues == null)
                ? OfficeOpenXmlChartWriter.ParseNumericCategories(data.Categories)
                : null;
            return new PowerPointScatterChartData(data.Series.Select(series =>
                new PowerPointScatterChartSeries(series.Name, series.XValues ?? sharedX!, series.Values)));
        }

    }
}
