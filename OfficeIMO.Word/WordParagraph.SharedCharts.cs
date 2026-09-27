using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Inserts an editable chart with an embedded worksheet after this paragraph.</summary>
    /// <param name="chartKind">Default chart family for series without their own render kind.</param>
    /// <param name="data">Shared chart categories, series, and appearance.</param>
    /// <param name="title">Optional chart title.</param>
    /// <param name="roundedCorners">Whether the chart frame has rounded corners.</param>
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
