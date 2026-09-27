using System.Collections.Generic;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

internal static partial class OfficeOpenXmlChartWriter {
    // Each data-source container has an exclusive choice. Replacing the complete
    // container also handles direct titles, numeric categories, and missing sources.
    internal static void UpdateCategorySeriesData(OpenXmlCompositeElement series, int index,
        string name, IReadOnlyList<string> categories, IReadOnlyList<double> values) {
        UpdateSeriesIndexOrder(series, index);
        string column = ColumnLetter(index + 2);
        series.AddChild(new C.SeriesText(CreateStringReference($"Sheet1!${column}$1", new[] { name })), true);
        series.AddChild(new C.CategoryAxisData(CreateStringReference($"Sheet1!$A$2:$A${categories.Count + 1}", categories)), true);
        series.AddChild(new C.Values(CreateNumberReference($"Sheet1!${column}$2:${column}${values.Count + 1}", values)), true);
    }
}
