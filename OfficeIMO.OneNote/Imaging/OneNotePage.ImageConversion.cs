using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

public static partial class OneNoteImageExportExtensions {
    /// <summary>Renders the selected page using the encoding in <see cref="OfficeImageExportOptions.Format"/>.</summary>
    public static byte[] ToImage(this OneNotePage document, OneNotePageRenderingOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToImage(options, out _, cancellationToken);

    /// <summary>Renders the selected page and returns its immutable fidelity report.</summary>
    public static byte[] ToImage(this OneNotePage document, OneNotePageRenderingOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        OneNotePageRenderingOptions resolved = SnapshotOneNotePageImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Export(cancellationToken);
        report = OfficeImageExportEvidence.CreateReport("OneNote", result, timer.Elapsed);
        return result.Bytes;
    }

    /// <summary>Renders the selected page with cancellation and returns the bytes and fidelity report.</summary>
    public static async Task<OfficeConversionResult<byte[]>> ToImageAsync(this OneNotePage document,
        OneNotePageRenderingOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        OneNotePageRenderingOptions resolved = SnapshotOneNotePageImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .ExportAsync(cancellationToken).ConfigureAwait(false);
        return new OfficeConversionResult<byte[]>(result.Bytes, OfficeImageExportEvidence.CreateReport("OneNote", result, timer.Elapsed));
    }

    /// <summary>Writes the selected page to a file and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this OneNotePage document, string path, OneNotePageRenderingOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        OneNotePageRenderingOptions resolved = SnapshotOneNotePageImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Save(path, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("OneNote", result, timer.Elapsed);
    }

    /// <summary>Writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this OneNotePage document, Stream stream, OneNotePageRenderingOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        OneNotePageRenderingOptions resolved = SnapshotOneNotePageImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Save(stream, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("OneNote", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a file and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this OneNotePage document, string path,
        OneNotePageRenderingOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        OneNotePageRenderingOptions resolved = SnapshotOneNotePageImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("OneNote", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this OneNotePage document, Stream stream,
        OneNotePageRenderingOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        OneNotePageRenderingOptions resolved = SnapshotOneNotePageImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("OneNote", result, timer.Elapsed);
    }

    private static OneNotePageRenderingOptions SnapshotOneNotePageImageOptions(OneNotePageRenderingOptions? options) =>
        options?.Clone() ?? new OneNotePageRenderingOptions();

    private static OfficeImageExportOperation<OneNotePageRenderingOptions> CreateImageExportOperation(
        OneNotePage document, OneNotePageRenderingOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return new OfficeImageExportOperation<OneNotePageRenderingOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => OneNotePageImageRenderer.Render(document, encoding, effective),
            exportWithCancellation: (encoding, effective, token) => OneNotePageImageRenderer.Render(document, encoding, effective, cancellationToken: token));
    }
}
