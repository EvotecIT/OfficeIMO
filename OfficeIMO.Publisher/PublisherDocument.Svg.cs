using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Exports one document page, selected by zero-based index, with combined source and rendering diagnostics.</summary>
    public OfficeConversionResult<string, PublisherConversionReport> ToSvgResult(int pageIndex = 0, PublisherSvgOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (pageIndex < 0 || pageIndex >= Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        var settings = options ?? new PublisherSvgOptions();
        var images = new List<OfficeImageExportDiagnostic>();
        var codec = new OfficeRasterImageFallbackCodec(settings.ImageCodec, images, "page/" + pageIndex);
        string svg = OfficeDrawingSvgExporter.ToSvg(Pages[pageIndex].Drawing, settings.Scale, settings.SizeUnit, codec, settings.ResourceIdPrefix, cancellationToken);
        return new OfficeConversionResult<string, PublisherConversionReport>(svg, new PublisherConversionReport(ReadReport, images));
    }
    /// <summary>Exports one document page as SVG. Use ToSvgResult to inspect source and rendering losses before accepting output.</summary>
    public string ToSvg(int pageIndex = 0, PublisherSvgOptions? options = null, CancellationToken cancellationToken = default) =>
        ToSvgResult(pageIndex, options, cancellationToken).Value;
}
