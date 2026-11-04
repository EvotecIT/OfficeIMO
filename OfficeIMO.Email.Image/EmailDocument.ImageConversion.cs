using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Email;

public static partial class EmailImageExportExtensions {
    /// <summary>Renders the selected page using the encoding in <see cref="OfficeImageExportOptions.Format"/>.</summary>
    public static byte[] ToImage(this EmailDocument document, EmailImageExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToImage(options, out _, cancellationToken);

    /// <summary>Renders the selected page and returns its immutable fidelity report.</summary>
    public static byte[] ToImage(this EmailDocument document, EmailImageExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EmailImageExportOptions resolved = SnapshotEmailDocumentImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Export(cancellationToken);
        report = OfficeImageExportEvidence.CreateReport("Email", result, timer.Elapsed);
        return result.Bytes;
    }

    /// <summary>Renders the selected page with cancellation and returns the bytes and fidelity report.</summary>
    public static async Task<OfficeConversionResult<byte[]>> ToImageAsync(this EmailDocument document,
        EmailImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EmailImageExportOptions resolved = SnapshotEmailDocumentImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .ExportAsync(cancellationToken).ConfigureAwait(false);
        return new OfficeConversionResult<byte[]>(result.Bytes, OfficeImageExportEvidence.CreateReport("Email", result, timer.Elapsed));
    }

    /// <summary>Writes the selected page to a file and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this EmailDocument document, string path, EmailImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EmailImageExportOptions resolved = SnapshotEmailDocumentImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Save(path, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Email", result, timer.Elapsed);
    }

    /// <summary>Writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this EmailDocument document, Stream stream, EmailImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EmailImageExportOptions resolved = SnapshotEmailDocumentImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Save(stream, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Email", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a file and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this EmailDocument document, string path,
        EmailImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EmailImageExportOptions resolved = SnapshotEmailDocumentImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Email", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this EmailDocument document, Stream stream,
        EmailImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EmailImageExportOptions resolved = SnapshotEmailDocumentImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Email", result, timer.Elapsed);
    }

    private static EmailImageExportOptions SnapshotEmailDocumentImageOptions(EmailImageExportOptions? options) =>
        options?.CloneEmail() ?? new EmailImageExportOptions();

    private static OfficeImageExportOperation<EmailImageExportOptions> CreateImageExportOperation(
        EmailDocument document, EmailImageExportOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return new OfficeImageExportOperation<EmailImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ExportImage(encoding, effective, effective.PageIndex),
            exportWithCancellation: (encoding, effective, token) => document.ExportImage(encoding, effective, effective.PageIndex, token),
            exportAsync: (encoding, effective, token) => document.ExportImageAsync(encoding, effective, effective.PageIndex, token));
    }
}
