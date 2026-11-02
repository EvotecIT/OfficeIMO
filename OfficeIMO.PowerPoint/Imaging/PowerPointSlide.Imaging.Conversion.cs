using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Renders the slide using the encoding in <see cref="OfficeImageExportOptions.Format"/>.</summary>
    public byte[] ToImage(PowerPointImageExportOptions? options = null,
        CancellationToken cancellationToken = default) => ToImage(options, out _, cancellationToken);

    /// <summary>Renders the slide and returns its immutable fidelity report.</summary>
    public byte[] ToImage(PowerPointImageExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(resolved, resolved.Format).Export(cancellationToken);
        report = OfficeImageExportEvidence.CreateReport("PowerPoint", result, timer.Elapsed);
        return result.Bytes;
    }

    /// <summary>Renders the slide with cancellation and returns the bytes and fidelity report.</summary>
    public async Task<OfficeConversionResult<byte[]>> ToImageAsync(PowerPointImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(resolved, resolved.Format)
            .ExportAsync(cancellationToken).ConfigureAwait(false);
        return new OfficeConversionResult<byte[]>(result.Bytes, OfficeImageExportEvidence.CreateReport("PowerPoint", result, timer.Elapsed));
    }

    /// <summary>Writes the slide to a file and returns its fidelity report.</summary>
    public OfficeConversionReport SaveAsImage(string path, PowerPointImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(resolved, resolved.Format).Save(path, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("PowerPoint", result, timer.Elapsed);
    }

    /// <summary>Writes the slide to a caller-owned stream and returns its fidelity report.</summary>
    public OfficeConversionReport SaveAsImage(Stream stream, PowerPointImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(resolved, resolved.Format).Save(stream, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("PowerPoint", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the slide to a file and returns its fidelity report.</summary>
    public async Task<OfficeConversionReport> SaveAsImageAsync(string path,
        PowerPointImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(resolved, resolved.Format)
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("PowerPoint", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the slide to a caller-owned stream and returns its fidelity report.</summary>
    public async Task<OfficeConversionReport> SaveAsImageAsync(Stream stream,
        PowerPointImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(resolved, resolved.Format)
            .SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("PowerPoint", result, timer.Elapsed);
    }

}
