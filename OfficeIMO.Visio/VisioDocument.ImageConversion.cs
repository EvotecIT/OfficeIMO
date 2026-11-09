using System.Diagnostics;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

public static partial class VisioImageExportExtensions {
    /// <summary>Renders the selected page using the encoding in <see cref="OfficeImageExportOptions.Format"/>.</summary>
    public static OfficeImageExportResult ToImage(this VisioDocument document, VisioImageExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToImage(options, out _, cancellationToken);

    /// <summary>Renders the selected page and returns its immutable fidelity report.</summary>
    public static OfficeImageExportResult ToImage(this VisioDocument document, VisioImageExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioDocumentImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Export(cancellationToken);
        report = OfficeImageExportEvidence.CreateReport("Visio", result, timer.Elapsed);
        return result;
    }

    /// <summary>Renders the selected page with cancellation and returns the encoded image and its diagnostics.</summary>
    public static async Task<OfficeImageExportResult> ToImageAsync(this VisioDocument document,
        VisioImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioDocumentImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .ExportAsync(cancellationToken).ConfigureAwait(false);
        return result;
    }

    /// <summary>Writes the selected page to a file and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this VisioDocument document, string path, VisioImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioDocumentImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Save(path, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Visio", result, timer.Elapsed);
    }

    /// <summary>Writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this VisioDocument document, Stream stream, VisioImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioDocumentImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Save(stream, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Visio", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a file and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this VisioDocument document, string path,
        VisioImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioDocumentImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Visio", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this VisioDocument document, Stream stream,
        VisioImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = SnapshotVisioDocumentImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Visio", result, timer.Elapsed);
    }

    private static VisioImageExportOptions SnapshotVisioDocumentImageOptions(VisioImageExportOptions? options) =>
        options?.Clone() ?? new VisioImageExportOptions();

    private static OfficeImageExportOperation<VisioImageExportOptions> CreateImageExportOperation(
        VisioDocument document, VisioImageExportOptions options) {
        ArgumentNullException.ThrowIfNull(document);
        return new OfficeImageExportOperation<VisioImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ToImage(encoding, effective),
            exportWithCancellation: (encoding, effective, token) => document.ToImage(encoding, effective, token));
    }
}
