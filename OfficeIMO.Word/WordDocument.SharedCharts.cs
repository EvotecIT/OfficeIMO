using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Creates an editable native chart and embedded worksheet from the shared OfficeIMO chart contract.</summary>
    /// <param name="chartKind">Default chart family, with optional per-series kinds for supported combinations.</param>
    /// <param name="data">Categories, series, and their supported appearance.</param>
    /// <param name="title">Optional chart title.</param>
    /// <param name="roundedCorners">Whether the native chart frame uses rounded corners.</param>
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
