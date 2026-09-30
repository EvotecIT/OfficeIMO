using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Creates an editable native chart and embedded worksheet from the shared OfficeIMO chart contract.</summary>
    /// <param name="chartKind">Default chart family, with optional per-series kinds for supported combinations.</param>
    /// <param name="data">Categories, series, and their supported appearance.</param>
    /// <param name="title">Optional chart title.</param>
    /// <param name="roundedCorners">Whether the native chart frame uses rounded corners.</param>
    /// <param name="width">Drawing width in pixels.</param>
    /// <param name="height">Drawing height in pixels.</param>
    /// <returns>The created chart.</returns>
    public WordChart AddChart(OfficeChartKind chartKind, OfficeChartData data, string title = "",
        bool roundedCorners = false, int width = 600, int height = 600) {
        if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        byte[] workbook = OfficeOpenXmlChartWriter.BuildWorkbook(data, chartKind);
        return AddChart(title, roundedCorners, width, height).ConfigureSharedData(chartKind, data, workbook);
    }
}
