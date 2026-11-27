using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

public partial class WordChart {
    private void RejectPieAppendToDoughnut() {
        OfficeXmlNode plot = PlotAreaOf(ResolveChart());
        if (!plot.IsNone && !plot.Child(NsC, "doughnutChart").IsNone)
            throw new NotSupportedException("Use AddDoughnut to append slices to a doughnut chart.");
    }

    /// <summary>Adds a category and value to an editable native doughnut chart with a 50-percent hole.</summary>
    /// <typeparam name="T">An int, double, or float value.</typeparam>
    /// <param name="category">The slice's category label.</param>
    /// <param name="value">A finite, nonnegative slice value.</param>
    /// <returns>The current chart.</returns>
    /// <exception cref="NotSupportedException">The chart already has another family, or its values use worksheet references.</exception>
    public WordChart AddDoughnut<T>(string category, T value) {
        if (!(value is int || value is double || value is float))
            throw new NotSupportedException("Value must be of type int, double, or float.");
        double number = OfficeIMO.Core.Internal.OfficeRoundTripText.ToDouble(value, CultureInfo.InvariantCulture);
        if (double.IsNaN(number) || double.IsInfinity(number) || number < 0)
            throw new ArgumentOutOfRangeException(nameof(value), "Doughnut values must be finite and nonnegative.");
        PrepareLiteralSliceAppend(category, doughnut: true);
        if (_chart.IsNone) {
            OfficeXmlNode chart = GenerateChart();
            OfficeXmlNode plotArea = PlotAreaOf(chart);
            OfficeXmlNode doughnut = NewC(plotArea, "doughnutChart");
            AddC(doughnut, "varyColors", true);
            doughnut.Append(AddDataLabel(doughnut));
            AddC(doughnut, "holeSize", "50");
            plotArea.Append(doughnut);
            AttachNewChart(chart);
            EnsureFramelessChartArea();
            UpdateTitle();
        }
        AddSingleCategory(category);
        AddSingleValue(value);
        return this;
    }
}
