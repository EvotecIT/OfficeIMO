using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Adds an editable native chart with the format's default title, size and placement.</summary>
    public WordChart AddChart(OfficeChartType chartKind, OfficeChartData data) =>
        AddChart(chartKind, data, null);

    /// <summary>Inserts an editable chart with an embedded worksheet after this paragraph.</summary>
    /// <param name="chartKind">Default chart family for series without their own render kind.</param>
    /// <param name="data">Shared chart categories, series, and appearance.</param>
    /// <param name="options">Optional title, physical size and Word frame options. The default size is 600 by 600 pixels.</param>
    /// <returns>The created chart.</returns>
    public WordChart AddChart(OfficeChartType chartKind, OfficeChartData data, OfficeChartOptions? options) {
        OfficeLength? width = options?.Width;
        OfficeLength? height = options?.Height;
        if ((width ?? OfficeLength.Pixels(600)).ToEmus() <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Chart width must be positive.");
        if ((height ?? OfficeLength.Pixels(600)).ToEmus() <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Chart height must be positive.");
        byte[] workbook = OfficeOpenXmlChartWriter.BuildWorkbook(data, chartKind);
        return AddChart(options?.Title ?? "", (options as WordChartOptions)?.RoundedCorners ?? false, width, height).ConfigureSharedData(chartKind, data, workbook);
    }
}
