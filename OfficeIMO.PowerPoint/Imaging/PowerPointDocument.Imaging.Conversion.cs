using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
#pragma warning disable RS0026 // Options select this overload; the explicit-encoding overload requires an enum first.
    /// <summary>Renders a batch using the encoding and selection in the options.</summary>
    public IReadOnlyList<OfficeImageExportResult> ToImages(PowerPointImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        PowerPointImageExportOptions resolved = NormalizeDocumentImageExportOptions(options);
        return CreateImageExportBatchOperation(resolved, resolved.Format).Export(cancellationToken);
    }

    /// <summary>Renders a batch asynchronously using the same options and evidence contract.</summary>
    public Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(PowerPointImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        PowerPointImageExportOptions resolved = NormalizeDocumentImageExportOptions(options);
        return CreateImageExportBatchOperation(resolved, resolved.Format).ExportAsync(cancellationToken);
    }

#pragma warning restore RS0026

    /// <summary>Saves the selected slides without retaining their encoded payloads, and returns an aggregate report.</summary>
    public OfficeConversionReport SaveAsImages(string folderPath, PowerPointImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PowerPointImageExportOptions resolved = NormalizeDocumentImageExportOptions(options);
        OfficeImageExportBatchSaveResult saved = CreateImageExportBatchOperation(resolved, resolved.Format).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("PowerPoint", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously saves the selected slides without retaining payloads, and returns an aggregate report.</summary>
    public async Task<OfficeConversionReport> SaveAsImagesAsync(string folderPath,
        PowerPointImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PowerPointImageExportOptions resolved = NormalizeDocumentImageExportOptions(options);
        OfficeImageExportBatchSaveResult saved = await CreateImageExportBatchOperation(resolved, resolved.Format)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("PowerPoint", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }
}
