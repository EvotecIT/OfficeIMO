using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.OpenDocument;

/// <summary>Thin ODP image-export bridge over the PowerPoint visual renderer.</summary>
public static partial class PowerPointOpenDocumentImageExportExtensions {
    /// <summary>Converts an ODP presentation to PowerPoint semantics and exports selected slides.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(
        this OdpPresentation source,
        OfficeImageExportFormat format,
        PowerPointImageExportOptions? imageOptions = null,
        PowerPointOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        source.ToImages(
            format,
            results.Add,
            imageOptions,
            conversionOptions,
            cancellationToken);
        return results.AsReadOnly();
    }

    /// <summary>Streams selected ODP slide images without retaining earlier payloads.</summary>
    public static void ToImages(
        this OdpPresentation source,
        OfficeImageExportFormat format,
        OfficeImageExportConsumer consumer,
        PowerPointImageExportOptions? imageOptions = null,
        PowerPointOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(consumer);
        PowerPointImageExportOptions effective =
            imageOptions?.Clone() ??
            new PowerPointImageExportOptions();
        OfficeImageExportExecutionScope.Run(effective, cancellationToken, token => {
            OdfConversionResult<PowerPointDocument> conversion =
                source.ConvertToPowerPointWithEvidence(conversionOptions);
            using (conversion.Value) {
                conversion.Value.ToImages(
                    format,
                    result => consumer(effective.EnsureAccepted(
                        OdfImageExportDiagnostics.Attach(result, conversion.Report))),
                    effective,
                    token);
            }
            return 0;
        });
    }
}
