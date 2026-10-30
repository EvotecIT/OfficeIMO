using OfficeIMO.Core.Internal;
using OfficeIMO.Publisher.Internal;
using System.Threading.Tasks;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Prepares bounded native PUB bytes and operation-level evidence. The returned array is caller-owned.</summary>
    /// <remarks>Supports exact source copying and the documented source-edit profiles. Source-free authoring and unsupported edits are refused even with Allow. Unassessed output requires explicit Allow. The document's original source baseline is unchanged.</remarks>
    public OfficeConversionResult<byte[], PublisherSaveReport> ToPubBytesResult(PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        PrepareSave((options ?? new PublisherSaveOptions()).Clone(), cancellationToken);

    /// <summary>Prepares native PUB bytes under the selected loss policy. Use ToPubBytesResult to retain operation evidence.</summary>
    public byte[] ToPubBytes(PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        ToPubBytesResult(options, cancellationToken).Value;

    /// <summary>Prepares a complete native PUB and atomically publishes it to an explicit .pub path.</summary>
    /// <remarks>Preparation, policy and pre-publication cancellation failures preserve an existing destination. Atomic replacement must be supported when Replace is selected. Saving does not rebind the document's source or destination.</remarks>
    public PublisherSaveReport Save(string path, PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) {
        string target = NativeSavePath(path);
        PublisherSaveOptions settings = (options ?? new PublisherSaveOptions()).Clone();
        var result = PrepareSave(settings, cancellationToken);
        OfficeFileCommit.WriteAtomically(target, destination => destination.Write(result.Value, 0, result.Value.Length),
            cancellationToken, FileConflict(settings));
        return result.Report;
    }

    /// <summary>Prepares native PUB bytes before writing a caller-owned stream, which remains open.</summary>
    /// <remarks>Seekable destinations are replaced and rewound after success; non-seekable destinations receive bytes at their current position. Preparation, policy and cancellation failures before publication preserve bytes and position. Stream I/O failures during publication may leave partial output. Synchronous stream cancellation is observed before publication.</remarks>
    public PublisherSaveReport Save(Stream stream, PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) {
        RequireNativeDestination(stream);
        var result = PrepareSave((options ?? new PublisherSaveOptions()).Clone(), cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeStreamWriter.WriteAllBytes(stream, result.Value);
        return result.Report;
    }

    /// <summary>Asynchronously stages and atomically publishes a complete native PUB to an explicit .pub path.</summary>
    /// <remarks>Preparation is performed in process before asynchronous I/O. Existing destinations survive pre-publication failures. Saving does not rebind the document's original source baseline.</remarks>
    public async Task<PublisherSaveReport> SaveAsync(string path, PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) {
        string target = NativeSavePath(path);
        PublisherSaveOptions settings = (options ?? new PublisherSaveOptions()).Clone();
        var result = PrepareSave(settings, cancellationToken);
        await OfficeFileCommit.WriteAtomicallyAsync(target, (destination, token) => destination.WriteAsync(result.Value, 0, result.Value.Length, token),
            cancellationToken, FileConflict(settings)).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Asynchronously writes a complete native PUB to a caller-owned stream without closing it.</summary>
    /// <remarks>Seekable destinations are replaced and rewound after success; non-seekable destinations receive bytes at their current position. Preparation failures leave bytes and position unchanged. Cancellation or I/O failure during stream publication can leave partial output.</remarks>
    public async Task<PublisherSaveReport> SaveAsync(Stream stream, PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) {
        RequireNativeDestination(stream);
        var result = PrepareSave((options ?? new PublisherSaveOptions()).Clone(), cancellationToken);
        await OfficeStreamWriter.WriteAllBytesAsync(stream, result.Value, cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    private OfficeConversionResult<byte[], PublisherSaveReport> PrepareSave(PublisherSaveOptions settings, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        PublisherNativePreparation prepared = PrepareNativeBytes(settings.MaxOutputBytes, token);
        var report = new PublisherSaveReport(SourceReadReport, prepared);
        if (!prepared.IsPrepared || report.HasErrors)
            throw new OfficeConversionException("The current publication cannot be saved through the supported native source profiles. Inspect the save report.", report);
        if (settings.LossPolicy == OfficeConversionLossPolicy.Block) report.RequireNoLoss();
        byte[] bytes = prepared.CopyBytes();
        token.ThrowIfCancellationRequested();
        return new OfficeConversionResult<byte[], PublisherSaveReport>(bytes, report);
    }

    private static string NativeSavePath(string path) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (string.IsNullOrWhiteSpace(path) || !string.Equals(Path.GetExtension(path), ".pub", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Native Publisher saving requires a .pub destination.", nameof(path));
        return Path.GetFullPath(path);
    }

    private static void RequireNativeDestination(Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The native PUB destination must be writable.", nameof(stream));
    }

    private static OfficeFileCommit.ConflictPolicy FileConflict(PublisherSaveOptions settings) =>
        settings.FileConflictPolicy == OfficeConversionFileConflictPolicy.Replace
            ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists;
}
