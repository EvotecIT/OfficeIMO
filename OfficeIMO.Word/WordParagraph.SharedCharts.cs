using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Inserts an editable chart with an embedded worksheet after this paragraph.</summary>
    /// <param name="chartKind">Default chart family for series without their own render kind.</param>
    /// <param name="data">Shared chart categories, series, and appearance.</param>
    /// <param name="title">Optional chart title.</param>
    /// <param name="roundedCorners">Whether the chart frame has rounded corners.</param>
    /// <param name="width">Drawing width; defaults to 600 pixels.</param>
    /// <param name="height">Drawing height; defaults to 600 pixels.</param>
    /// <returns>The created chart.</returns>
    public WordChart AddChart(OfficeChartType chartKind, OfficeChartData data, string title = "",
        bool roundedCorners = false, OfficeLength? width = null, OfficeLength? height = null) {
        if ((width ?? OfficeLength.Pixels(600)).ToEmus() <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if ((height ?? OfficeLength.Pixels(600)).ToEmus() <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        byte[] workbook = OfficeOpenXmlChartWriter.BuildWorkbook(data, chartKind);
        return AddChart(title, roundedCorners, width, height).ConfigureSharedData(chartKind, data, workbook);
    }
}
