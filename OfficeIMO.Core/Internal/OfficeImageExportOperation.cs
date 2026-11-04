using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing;

/// <summary>
/// Executes one image export from options, keeping rendering policy, deadlines and persistence
/// independent of the public authoring surface. Adapters share this owner through friend access.
/// </summary>
internal sealed class OfficeImageExportOperation<TOptions> where TOptions : OfficeImageExportOptions {
    private readonly TOptions _options;
    private readonly OfficeImageExportFormat _format;
    private readonly OfficeImageExportFileConflictPolicy _conflictPolicy;
    private readonly Func<OfficeImageExportFormat, TOptions, OfficeImageExportResult> _export;
    private readonly Func<OfficeImageExportFormat, TOptions, CancellationToken, OfficeImageExportResult>? _exportWithCancellation;
    private readonly Func<OfficeImageExportFormat, TOptions, CancellationToken, Task<OfficeImageExportResult>>? _exportAsync;

    internal OfficeImageExportOperation(
        TOptions options,
        OfficeImageExportFormat format,
        OfficeImageExportFileConflictPolicy conflictPolicy,
        Func<OfficeImageExportFormat, TOptions, OfficeImageExportResult> export,
        Func<OfficeImageExportFormat, TOptions, CancellationToken, OfficeImageExportResult>? exportWithCancellation = null,
        Func<OfficeImageExportFormat, TOptions, CancellationToken, Task<OfficeImageExportResult>>? exportAsync = null) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        if (!Enum.IsDefined(typeof(OfficeImageExportFormat), format)) throw new ArgumentOutOfRangeException(nameof(format));
        _format = format;
        _conflictPolicy = conflictPolicy;
        _export = export ?? throw new ArgumentNullException(nameof(export));
        _exportWithCancellation = exportWithCancellation;
        _exportAsync = exportAsync;
    }

    /// <summary>Exports using the currently configured format and options.</summary>
    internal OfficeImageExportResult Export() => Export(CancellationToken.None);

    /// <summary>Exports using the currently configured format and observes cancellation during supported render stages.</summary>
    internal OfficeImageExportResult Export(CancellationToken cancellationToken) {
        OfficeImageExportResult result = Render(cancellationToken);
        _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name));
        return result;
    }

    /// <summary>Exports asynchronously when the adapter owns a genuine asynchronous render path.</summary>
    internal async Task<OfficeImageExportResult> ExportAsync(CancellationToken cancellationToken = default) {
        OfficeImageExportResult result = await RenderAsync(cancellationToken).ConfigureAwait(false);
        _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name));
        return result;
    }

    /// <summary>Exports and returns the encoded image bytes.</summary>
    internal byte[] ToBytes() => Export().Bytes;

    /// <summary>Saves the exported image to a file path.</summary>
    internal OfficeImageExportResult Save(string path, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));
        }

        cancellationToken.ThrowIfCancellationRequested();
        string resolvedPath = OfficeImageExportPath.NormalizeFile(path, _format);
        OfficeImageExportResult result = Render(cancellationToken);
        _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, 0, 1, result.Name, resolvedPath));
        cancellationToken.ThrowIfCancellationRequested();
        OfficeImageExportResult saved = result.Save(path, _conflictPolicy);
        _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name, saved.SavedPath));
        return saved;
    }

    /// <summary>Writes the exported image to a stream.</summary>
    internal OfficeImageExportResult Save(Stream stream, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeImageExportResult result = Render(cancellationToken);
        _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, 0, 1, result.Name));
        cancellationToken.ThrowIfCancellationRequested();
        OfficeImageExportResult saved = result.Save(stream);
        _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name));
        return saved;
    }

    /// <summary>Asynchronously saves the exported image to a file path.</summary>
    internal async Task<OfficeImageExportResult> SaveAsync(
        string path,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));
        }

        cancellationToken.ThrowIfCancellationRequested();
        string resolvedPath = OfficeImageExportPath.NormalizeFile(path, _format);
        OfficeImageExportResult result = await RenderAsync(cancellationToken).ConfigureAwait(false);
        _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, 0, 1, result.Name, resolvedPath));
        OfficeImageExportResult saved = await result.SaveAsync(
            path,
            _conflictPolicy,
            cancellationToken).ConfigureAwait(false);
        _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name, saved.SavedPath));
        return saved;
    }

    /// <summary>Asynchronously writes the exported image to a stream.</summary>
    internal async Task<OfficeImageExportResult> SaveAsync(
        Stream stream,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeImageExportResult result = await RenderAsync(cancellationToken).ConfigureAwait(false);
        _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, 0, 1, result.Name));
        OfficeImageExportResult saved = await result.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name));
        return saved;
    }

    private OfficeImageExportResult Render(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        TOptions effective = _options.CreateEffectiveImageExportOptions<TOptions>();
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            effective.RenderTimeout,
            cancellationToken);
        try {
            effective.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Rendering, 0, 1));
            OfficeImageExportResult result = _exportWithCancellation != null
                ? _exportWithCancellation(_format, effective, execution.Token)
                : _export(_format, effective);
            execution.ThrowIfCancellationRequested();
            effective.EnsureAccepted(result);
            return result;
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    private async Task<OfficeImageExportResult> RenderAsync(CancellationToken cancellationToken) {
        if (_exportAsync == null) return Render(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        TOptions effective = _options.CreateEffectiveImageExportOptions<TOptions>();
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            effective.RenderTimeout,
            cancellationToken);
        try {
            effective.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Rendering, 0, 1));
            OfficeImageExportResult result = await _exportAsync(_format, effective, execution.Token).ConfigureAwait(false);
            execution.ThrowIfCancellationRequested();
            effective.EnsureAccepted(result);
            return result;
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

}
