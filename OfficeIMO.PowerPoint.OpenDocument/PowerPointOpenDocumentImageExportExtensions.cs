using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using System.Threading;

namespace OfficeIMO.PowerPoint.OpenDocument;

/// <summary>Thin ODP image-export bridge over the PowerPoint visual renderer.</summary>
public static class PowerPointOpenDocumentImageExportExtensions {
    /// <summary>Converts an ODP presentation to PowerPoint semantics and exports selected slides.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this OdpPresentation source,
        OfficeImageExportFormat format,
        PowerPointPresentationImageExportOptions? imageOptions = null,
        PowerPointOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        source.ExportImages(
            format,
            results.Add,
            imageOptions,
            conversionOptions,
            cancellationToken);
        return results.AsReadOnly();
    }

    /// <summary>Streams selected ODP slide images without retaining earlier payloads.</summary>
    public static void ExportImages(
        this OdpPresentation source,
        OfficeImageExportFormat format,
        OfficeImageExportConsumer consumer,
        PowerPointPresentationImageExportOptions? imageOptions = null,
        PowerPointOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        PowerPointPresentationImageExportOptions effective =
            imageOptions?.ClonePresentation() ??
            new PowerPointPresentationImageExportOptions();
        OdfConversionResult<PowerPointPresentation> conversion =
            source.ToPowerPointPresentationResult(conversionOptions);
        using (conversion.Value) {
            conversion.Value.ExportImages(
                format,
                result => consumer(effective.EnsureAccepted(
                    OdfImageExportDiagnostics.Attach(result, conversion.Report))),
                effective,
                cancellationToken);
        }
    }
}
