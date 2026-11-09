using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.OpenDocument;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentImageExportExtensions {
    /// <summary>Renders selected source surfaces using the encoding in the image options.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this OdpPresentation document,
        PowerPointImageExportOptions? imageOptions = null, PowerPointOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) =>
        CreateImageOperation(document, imageOptions?.Clone() ?? new PowerPointImageExportOptions(), conversionOptions).Export(cancellationToken);

    /// <summary>Renders selected source surfaces asynchronously under one shared operation deadline and budget.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(this OdpPresentation document,
        PowerPointImageExportOptions? imageOptions = null, PowerPointOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        var resolved = imageOptions?.Clone() ?? new PowerPointImageExportOptions();
        return CreateImageOperation(document, resolved, conversionOptions).ExportAsync(cancellationToken);
    }

    /// <summary>Saves selected source surfaces through the shared atomic file and fidelity policies.</summary>
    public static OfficeConversionReport SaveAsImages(this OdpPresentation document, string folderPath,
        PowerPointImageExportOptions? imageOptions = null, PowerPointOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        var resolved = imageOptions?.Clone() ?? new PowerPointImageExportOptions();
        var saved = CreateImageOperation(document, resolved, conversionOptions).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("ODP", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Saves selected source surfaces asynchronously without retaining earlier encoded payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this OdpPresentation document, string folderPath,
        PowerPointImageExportOptions? imageOptions = null, PowerPointOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        var resolved = imageOptions?.Clone() ?? new PowerPointImageExportOptions();
        var saved = await CreateImageOperation(document, resolved, conversionOptions)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("ODP", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<PowerPointImageExportOptions> CreateImageOperation(
        OdpPresentation document, PowerPointImageExportOptions options, PowerPointOpenDocumentConversionOptions? conversionOptions) {
        ArgumentNullException.ThrowIfNull(document);
        PowerPointOpenDocumentConversionOptions? conversion = conversionOptions?.Clone();
        return new OfficeImageExportBatchOperation<PowerPointImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (format, effective) => document.ToImages(format, effective, conversion),
            exportEach: (format, effective, consumer, token) => document.ToImages(format, consumer, effective, conversion, token));
    }
}
