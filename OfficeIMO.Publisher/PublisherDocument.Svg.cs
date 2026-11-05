using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Exports one document page, selected by zero-based index, with combined source and rendering diagnostics.</summary>
    public string ToSvg(int pageIndex, PublisherSvgOptions? options, out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentOutOfRangeException.ThrowIfNegative(pageIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(pageIndex, Pages.Count);
        var settings = options ?? new PublisherSvgOptions();
        var images = new List<OfficeImageExportDiagnostic>();
        var codec = new OfficeRasterImageFallbackCodec(settings.ImageCodec, images, "page/" + pageIndex);
        string svg = OfficeDrawingSvgExporter.ToSvg(Pages[pageIndex].CreateDrawing(), settings.Scale, settings.SizeUnit, codec, settings.ResourceIdPrefix, cancellationToken);
        report = new PublisherConversionReport(ReadReport, images);
        return svg;
    }
    /// <summary>Exports one document page as SVG. Use the report overload to inspect source and rendering losses before accepting output.</summary>
    public string ToSvg(int pageIndex = 0, PublisherSvgOptions? options = null, CancellationToken cancellationToken = default) =>
        ToSvg(pageIndex, options, out _, cancellationToken);
}
