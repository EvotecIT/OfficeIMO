using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown;

namespace OfficeIMO.Notebook.Markdown;

internal sealed partial class NotebookMarkdownProjection {
    private const string PlotlyMime = "application/vnd.plotly.v1+json";
    private static readonly string[] VegaLiteMimes = {
        "application/vnd.vegalite.v6+json", "application/vnd.vegalite.v6.json",
        "application/vnd.vegalite.v5+json", "application/vnd.vegalite.v5.json",
        "application/vnd.vegalite.v4+json", "application/vnd.vegalite.v4.json"
    };

    private bool TryAddChart(IReadOnlyDictionary<string, JsonElement> data, string location) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (data.TryGetValue(PlotlyMime, out JsonElement plotly)
            && TryAddChart(NotebookPlotlyChart.TryRead(plotly, _cancellationToken), "Plotly", location)) return true;
        foreach (string mime in VegaLiteMimes) {
            if (data.TryGetValue(mime, out JsonElement vegaLite)
                && TryAddChart(NotebookVegaLiteChart.TryRead(vegaLite, _cancellationToken), "Vega-Lite", location)) return true;
        }
        return false;
    }

    private bool TryAddChart(OfficeChartSnapshot? chart, string producer, string location) {
        if (chart == null) return false;
        _cancellationToken.ThrowIfCancellationRequested();
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(chart);
        byte[] svgBytes = OfficeDrawingSvgExporter.ToSvgBytes(drawing);
        if (svgBytes.Length > _options.MaxImageBytes) return false;
        _cancellationToken.ThrowIfCancellationRequested();
        string imageUri = "data:image/svg+xml;base64," + Convert.ToBase64String(svgBytes);
        _blocks.Add(new ImageBlock(imageUri, chart.Title ?? (producer + " chart"), null));
        AddUnsupported("notebook-output-chart-static-preview", location,
            "A static preview was drawn from the saved " + producer + " chart specification.",
            OfficeConversionLossKind.Approximation);
        return true;
    }
}
