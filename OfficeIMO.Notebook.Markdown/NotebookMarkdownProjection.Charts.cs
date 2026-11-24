using System.Globalization;
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

    private NotebookChartPreview? CreateChartPreview(IReadOnlyDictionary<string, JsonElement> data, bool interactive) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (data.TryGetValue(PlotlyMime, out JsonElement plotly)) {
            var marks = interactive ? new List<NotebookChartMark>() : null;
            if (NotebookPlotlyChart.TryRead(plotly, _cancellationToken) is OfficeChartSnapshot snapshot) {
                if (CreateChartPreview(RenderChartSnapshot(snapshot, marks), snapshot.Title, "Plotly", marks) is { } preview)
                    return preview;
            } else if (RenderPlotlyDrawing(plotly, marks, out string? title, out string producer) is { } drawing
                && CreateChartPreview(drawing, title, producer, marks) is { } preview) {
                return preview;
            }
        }
        foreach (string mime in VegaLiteMimes) {
            if (data.TryGetValue(mime, out JsonElement vegaLite)
                && NotebookVegaLiteChart.TryRead(vegaLite, _cancellationToken) is OfficeChartSnapshot snapshot) {
                var marks = interactive ? new List<NotebookChartMark>() : null;
                if (CreateChartPreview(RenderChartSnapshot(snapshot, marks), snapshot.Title, "Vega-Lite", marks) is { } preview)
                    return preview;
            }
        }
        return null;
    }

    private OfficeDrawing? RenderPlotlyDrawing(JsonElement spec, List<NotebookChartMark>? marks, out string? title,
        out string producer) {
        producer = "Plotly";
        if (NotebookPlotlyChart.TryRenderCartesianPanels(spec, _cancellationToken, out title, marks) is { } panels) return panels;
        marks?.Clear();
        if (NotebookPlotlyChart.TryRenderThreeDimensional(spec, _cancellationToken, out title, marks) is { } scene) {
            producer = "Plotly 3D";
            return scene;
        }
        marks?.Clear();
        return NotebookPlotlyChart.TryRenderGeographic(spec, _cancellationToken, out title, marks);
    }

    /// <summary>Captures matching point placements in the notebook canvas when rendering chart panels.</summary>
    internal static OfficeDrawing RenderChartSnapshot(OfficeChartSnapshot snapshot, List<NotebookChartMark>? marks,
        bool useMinimumCanvas = true, double offsetX = 0, double offsetY = 0) {
        if (marks == null) return OfficeChartDrawingRenderer.Render(snapshot, useMinimumCanvas);
        return OfficeChartDrawingRenderer.Render(snapshot, useMinimumCanvas, diagnostics: null, pointRendered: point => {
            string value = point.Value.ToString("G", CultureInfo.InvariantCulture);
            string label = point.NumericX is double x
                ? point.Series + " • x: " + x.ToString("G", CultureInfo.InvariantCulture) + " • y: " + value
                : point.Series + " • " + point.Category + ": " + value;
            marks.Add(new NotebookChartMark(offsetX + point.X, offsetY + point.Y, point.Width, point.Height, label));
        });
    }

    private NotebookChartPreview? CreateChartPreview(OfficeDrawing drawing, string? title, string producer,
        List<NotebookChartMark>? marks) {
        _cancellationToken.ThrowIfCancellationRequested();
        byte[] svgBytes;
        try {
            svgBytes = OfficeDrawingSvgExporter.ToSvgBytes(drawing, 1D, OfficeSvgSizeUnit.Point,
                imageCodec: null, resourceIdPrefix: null, maximumUtf8Bytes: _options.MaxImageBytes, _cancellationToken);
        } catch (OfficeImageExportBatchLimitException) {
            return null;
        }
        string imageUri = "data:image/svg+xml;base64," + Convert.ToBase64String(svgBytes);
        return new NotebookChartPreview(imageUri, drawing.Width, drawing.Height, title, producer, marks);
    }

    private void ReportChartPreview(NotebookChartPreview chart, string location, bool interactive) {
        string kind = interactive ? "An interactive approximation" : "A static preview";
        AddUnsupported(interactive ? "notebook-output-chart-interactive-preview" : "notebook-output-chart-static-preview",
            location, kind + " was drawn from the saved " + chart.Producer + " chart specification.",
            OfficeConversionLossKind.Approximation);
    }
}
