#pragma warning disable RS0026 // The required path or stream selects the SaveAsImage destination.
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointChart {
    /// <summary>Renders this chart on its own, without the rest of the slide.</summary>
    /// <remarks>Scheme colors and series without explicit colors resolve against the theme that applies to the chart's slide.</remarks>
    /// <param name="options">Image format, size and rendering settings; null renders a PNG at the chart's frame size.</param>
    /// <param name="cancellationToken">Token used to cancel rendering.</param>
    /// <returns>The encoded image, its size and the rendering diagnostics.</returns>
    /// <exception cref="NotSupportedException">The shared chart renderer cannot draw this chart; the message gives the reason.</exception>
    public OfficeImageExportResult ToImage(OfficeChartImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeChartImageExporter.ToImage(CreateImageSource(), options, cancellationToken);

    /// <summary>Renders this chart and writes the image to a file.</summary>
    /// <param name="path">Destination file path.</param>
    /// <param name="options">Image format, size and rendering settings; null renders a PNG at the chart's frame size.</param>
    /// <param name="cancellationToken">Token used to cancel rendering.</param>
    /// <returns>The conversion report with the rendering diagnostics.</returns>
    /// <exception cref="NotSupportedException">The shared chart renderer cannot draw this chart; the message gives the reason.</exception>
    public OfficeConversionReport SaveAsImage(string path, OfficeChartImageExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        OfficeChartImageExporter.SaveAsImage(CreateImageSource(), path, options, cancellationToken);

    /// <summary>Renders this chart and writes the image to a caller-owned stream.</summary>
    /// <param name="stream">Destination stream.</param>
    /// <param name="options">Image format, size and rendering settings; null renders a PNG at the chart's frame size.</param>
    /// <param name="cancellationToken">Token used to cancel rendering.</param>
    /// <returns>The conversion report with the rendering diagnostics.</returns>
    /// <exception cref="NotSupportedException">The shared chart renderer cannot draw this chart; the message gives the reason.</exception>
    public OfficeConversionReport SaveAsImage(Stream stream, OfficeChartImageExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        OfficeChartImageExporter.SaveAsImage(CreateImageSource(), stream, options, cancellationToken);

    private OfficeChartImageSource CreateImageSource() {
        string name = Name is { Length: > 0 } shapeName ? shapeName : "Chart";
        string source = "PowerPoint chart '" + name + "'";
        return new OfficeChartImageSource("PowerPoint", name, source, Width, Height, (width, height, fidelity, diagnostics) => {
            using var scope = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartFidelityScope.Enter(fidelity);
            if (!TryGetOfficeSnapshot(width, height, out OfficeChartSnapshot snapshot, out string? reason, out string? projectionNotice)) {
                throw new NotSupportedException(OfficeChartImageExporter.UnsupportedChartMessage + reason);
            }
            foreach (string approximation in scope.Approximations)
                diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                    OfficeDiagnosticCodes.ChartApproximation, approximation, source, OfficeConversionLossKind.Approximation));
            if (projectionNotice != null) {
                diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                    PowerPointImageExportDiagnosticCodes.ChartProjected, projectionNotice, source,
                    OfficeConversionLossKind.Approximation));
            }
            return snapshot;
        });
    }
}
