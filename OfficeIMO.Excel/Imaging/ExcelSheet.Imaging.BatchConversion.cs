using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
#pragma warning disable RS0026 // Options select this overload; the explicit-encoding overload requires an enum first.
    /// <summary>Renders a batch using the encoding and selection in the options.</summary>
    public IReadOnlyList<OfficeImageExportResult> ToImages(ExcelImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ExcelImageExportOptions resolved = NormalizeWorksheetOptions(options);
        return CreateImageExportBatchOperation(resolved, resolved.Format).Export(cancellationToken);
    }

    /// <summary>Renders a batch asynchronously using the same options and evidence contract.</summary>
    public Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(ExcelImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ExcelImageExportOptions resolved = NormalizeWorksheetOptions(options);
        return CreateImageExportBatchOperation(resolved, resolved.Format).ExportAsync(cancellationToken);
    }

#pragma warning restore RS0026

    /// <summary>Saves the selected worksheet images without retaining their encoded payloads, and returns an aggregate report.</summary>
    public OfficeConversionReport SaveAsImages(string folderPath, ExcelImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        ExcelImageExportOptions resolved = NormalizeWorksheetOptions(options);
        OfficeImageExportBatchSaveResult saved = CreateImageExportBatchOperation(resolved, resolved.Format).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Excel", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously saves the selected worksheet images without retaining payloads, and returns an aggregate report.</summary>
    public async Task<OfficeConversionReport> SaveAsImagesAsync(string folderPath,
        ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        ExcelImageExportOptions resolved = NormalizeWorksheetOptions(options);
        OfficeImageExportBatchSaveResult saved = await CreateImageExportBatchOperation(resolved, resolved.Format)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Excel", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }
}
