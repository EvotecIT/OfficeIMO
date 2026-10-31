using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

public static partial class VisioImageExportExtensions {
    /// <summary>Renders the selected page using the encoding in <see cref="OfficeImageExportOptions.Format"/>.</summary>
    public static byte[] ToImage(this VisioPage document, VisioImageExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToImage(options, out _, cancellationToken);

    /// <summary>Renders the selected page and returns its immutable fidelity report.</summary>
    public static byte[] ToImage(this VisioPage document, VisioImageExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioPageImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Export(cancellationToken);
        report = OfficeImageExportEvidence.CreateReport("Visio", result, timer.Elapsed);
        return result.Bytes;
    }

    /// <summary>Renders the selected page with cancellation and returns the bytes and fidelity report.</summary>
    public static async Task<OfficeConversionResult<byte[]>> ToImageAsync(this VisioPage document,
        VisioImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioPageImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .ExportAsync(cancellationToken).ConfigureAwait(false);
        return new OfficeConversionResult<byte[]>(result.Bytes, OfficeImageExportEvidence.CreateReport("Visio", result, timer.Elapsed));
    }

    /// <summary>Writes the selected page to a file and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this VisioPage document, string path, VisioImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioPageImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Save(path, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Visio", result, timer.Elapsed);
    }

    /// <summary>Writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this VisioPage document, Stream stream, VisioImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioPageImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Save(stream, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Visio", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a file and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this VisioPage document, string path,
        VisioImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioPageImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Visio", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this VisioPage document, Stream stream,
        VisioImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioPageImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Visio", result, timer.Elapsed);
    }

    private static VisioImageExportOptions SnapshotVisioPageImageOptions(VisioImageExportOptions? options) =>
        options?.Clone() ?? new VisioImageExportOptions();

    private static OfficeImageExportOperation<VisioImageExportOptions> CreateImageExportOperation(
        VisioPage document, VisioImageExportOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return new OfficeImageExportOperation<VisioImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ExportImage(encoding, effective),
            exportWithCancellation: (encoding, effective, token) => document.ExportImage(encoding, effective, token));
    }
}
