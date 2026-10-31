using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word;

public static partial class WordImageExportExtensions {
    /// <summary>Renders the selected page using the encoding in <see cref="OfficeImageExportOptions.Format"/>.</summary>
    public static byte[] ToImage(this WordDocument document, WordImageExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToImage(options, out _, cancellationToken);

    /// <summary>Renders the selected page and returns its immutable fidelity report.</summary>
    public static byte[] ToImage(this WordDocument document, WordImageExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        WordImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved, resolved.Format).Export(cancellationToken);
        report = OfficeImageExportEvidence.CreateReport("Word", result, timer.Elapsed);
        return result.Bytes;
    }

    /// <summary>Renders the selected page with cancellation and returns the bytes and fidelity report.</summary>
    public static async Task<OfficeConversionResult<byte[]>> ToImageAsync(this WordDocument document,
        WordImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        WordImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved, resolved.Format)
            .ExportAsync(cancellationToken).ConfigureAwait(false);
        return new OfficeConversionResult<byte[]>(result.Bytes, OfficeImageExportEvidence.CreateReport("Word", result, timer.Elapsed));
    }

    /// <summary>Writes the selected page to a file and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this WordDocument document, string path, WordImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        WordImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved, resolved.Format).Save(path, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Word", result, timer.Elapsed);
    }

    /// <summary>Writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this WordDocument document, Stream stream, WordImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        WordImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved, resolved.Format).Save(stream, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Word", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a file and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this WordDocument document, string path,
        WordImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        WordImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved, resolved.Format)
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Word", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this WordDocument document, Stream stream,
        WordImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        WordImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved, resolved.Format)
            .SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Word", result, timer.Elapsed);
    }

    /// <summary>Saves the selected pages without retaining their encoded payloads, and returns an aggregate report.</summary>
    public static OfficeConversionReport SaveAsImages(this WordDocument document, string folderPath, WordImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        WordImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportBatchSaveResult saved = CreateImageExportBatchOperation(document, resolved, resolved.Format).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Word", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously saves the selected pages without retaining payloads, and returns an aggregate report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this WordDocument document, string folderPath,
        WordImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        WordImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportBatchSaveResult saved = await CreateImageExportBatchOperation(document, resolved, resolved.Format)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Word", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }
}
