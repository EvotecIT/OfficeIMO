using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal;

internal static partial class OfficeOpenXmlChartWriter {
    /// <summary>A newly added series has no source point overrides, custom labels or measured error magnitudes.</summary>
    internal static void ClearAddedSeriesMetadata(OfficeXmlNode series) {
        series.RemoveChildren(NsC, "trendline");
        series.RemoveChildren(NsC, "errBars");
        series.RemoveChildren(NsC, "dPt");
        series.RemoveChildren(NsC, "dLbls");
    }
    // Each data-source container has an exclusive choice. Replacing the complete
    // container also handles direct titles, numeric categories, and missing sources.
    internal static void UpdateCategorySeriesData(OfficeXmlNode series, int index,
        string name, IReadOnlyList<string> categories, IReadOnlyList<double> values) {
        UpdateSeriesIndexOrder(series, index);
        string column = ColumnLetter(index + 2);
        OfficeXmlChildOrder order = SeriesOrder(SeriesLayer(series, null))
            ?? throw new InvalidOperationException("The series is not under a category chart layer.");
        OfficeXmlNode text = series.CreateElement(NsC, "tx");
        text.Append(CreateStringReference(text, $"Sheet1!${column}$1", new[] { name }));
        SetOrdered(series, text, order);
        OfficeXmlNode categoryData = series.CreateElement(NsC, "cat");
        categoryData.Append(CreateStringReference(categoryData, $"Sheet1!$A$2:$A${categories.Count + 1}", categories));
        SetOrdered(series, categoryData, order);
        OfficeXmlNode valueData = series.CreateElement(NsC, "val");
        valueData.Append(CreateNumberReference(valueData, $"Sheet1!${column}$2:${column}${values.Count + 1}", values));
        SetOrdered(series, valueData, order);
    }

    // SDK callers

    internal static void ClearAddedSeriesMetadata(OpenXmlCompositeElement series) =>
        ClearAddedSeriesMetadata(OfficeXmlNode.FromSdk(series));

    internal static void UpdateCategorySeriesData(OpenXmlCompositeElement series, int index,
        string name, IReadOnlyList<string> categories, IReadOnlyList<double> values) =>
        UpdateCategorySeriesData(OfficeXmlNode.FromSdk(series), index, name, categories, values);
}
