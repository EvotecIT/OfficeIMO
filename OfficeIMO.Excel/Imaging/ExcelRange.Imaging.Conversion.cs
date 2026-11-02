using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel;

public sealed partial class ExcelRange {
    /// <summary>Renders the range using the encoding in <see cref="OfficeImageExportOptions.Format"/>.</summary>
    public byte[] ToImage(ExcelImageExportOptions? options = null,
        CancellationToken cancellationToken = default) => ToImage(options, out _, cancellationToken);

    /// <summary>Renders the range and returns its immutable fidelity report.</summary>
    public byte[] ToImage(ExcelImageExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        ExcelImageExportOptions resolved = NormalizeOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(resolved, resolved.Format).Export(cancellationToken);
        report = OfficeImageExportEvidence.CreateReport("Excel", result, timer.Elapsed);
        return result.Bytes;
    }

    /// <summary>Renders the range with cancellation and returns the bytes and fidelity report.</summary>
    public async Task<OfficeConversionResult<byte[]>> ToImageAsync(ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        ExcelImageExportOptions resolved = NormalizeOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(resolved, resolved.Format)
            .ExportAsync(cancellationToken).ConfigureAwait(false);
        return new OfficeConversionResult<byte[]>(result.Bytes, OfficeImageExportEvidence.CreateReport("Excel", result, timer.Elapsed));
    }

    /// <summary>Writes the range to a file and returns its fidelity report.</summary>
    public OfficeConversionReport SaveAsImage(string path, ExcelImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        ExcelImageExportOptions resolved = NormalizeOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(resolved, resolved.Format).Save(path, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Excel", result, timer.Elapsed);
    }

    /// <summary>Writes the range to a caller-owned stream and returns its fidelity report.</summary>
    public OfficeConversionReport SaveAsImage(Stream stream, ExcelImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        ExcelImageExportOptions resolved = NormalizeOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(resolved, resolved.Format).Save(stream, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Excel", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the range to a file and returns its fidelity report.</summary>
    public async Task<OfficeConversionReport> SaveAsImageAsync(string path,
        ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        ExcelImageExportOptions resolved = NormalizeOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(resolved, resolved.Format)
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Excel", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the range to a caller-owned stream and returns its fidelity report.</summary>
    public async Task<OfficeConversionReport> SaveAsImageAsync(Stream stream,
        ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        ExcelImageExportOptions resolved = NormalizeOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(resolved, resolved.Format)
            .SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Excel", result, timer.Elapsed);
    }

}
