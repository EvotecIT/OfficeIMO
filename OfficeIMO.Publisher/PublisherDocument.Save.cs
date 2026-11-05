using OfficeIMO.Core.Internal;
using OfficeIMO.Publisher.Internal;
using System.Diagnostics;
using System.Threading.Tasks;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Prepares caller-owned native PUB bytes under the selected loss policy.</summary>
    public byte[] ToBytes(PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        ToBytes(options, out _, cancellationToken);

    /// <summary>Prepares caller-owned native PUB bytes and returns the save report.</summary>
    /// <remarks>Supports exact source copying and the documented source-edit profiles. Source-free publications and unsupported edits always refuse native saving. Unassessed output requires <see cref="PublisherSaveOptions.AllowLoss"/>. Preparing bytes leaves the original source baseline intact.</remarks>
    public byte[] ToBytes(PublisherSaveOptions? options, out PublisherSaveReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        byte[] bytes = PrepareSave((options ?? new PublisherSaveOptions()).Clone(), timer, cancellationToken, out var prepared);
        report = new PublisherSaveReport(SourceReadReport, prepared, timer.Elapsed, itemsWritten: 1);
        return bytes;
    }

    /// <summary>Atomically saves to the associated native path, replacing that file by default. Explicit settings retain their selected conflict policy.</summary>
    /// <remarks>A successful path load or native path save establishes the destination. Source-free publications still refuse native saving. The original source baseline remains unchanged.</remarks>
    public PublisherSaveReport Save(PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        Save(FilePath ?? throw new InvalidOperationException("The publication has no associated native path. Use Save(string) or Save(Stream)."),
            options ?? new PublisherSaveOptions { FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace }, cancellationToken);

    /// <summary>Prepares a complete native PUB and atomically publishes it to an explicit .pub path.</summary>
    /// <remarks>Preparation, policy and pre-publication cancellation failures preserve an existing destination and the associated path. Atomic replacement must be supported when Replace is selected. Success associates the destination without changing the original source baseline.</remarks>
    public PublisherSaveReport Save(string path, PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) {
        string target = NativeSavePath(path);
        var timer = Stopwatch.StartNew();
        PublisherSaveOptions settings = (options ?? new PublisherSaveOptions()).Clone();
        byte[] bytes = PrepareSave(settings, timer, cancellationToken, out var prepared);
        OfficeFileCommit.WriteAtomically(target, destination => destination.Write(bytes, 0, bytes.Length),
            cancellationToken, FileConflict(settings));
        FilePath = target;
        return new PublisherSaveReport(SourceReadReport, prepared, timer.Elapsed, itemsWritten: 1, outputPath: target);
    }

    /// <summary>Prepares native PUB bytes before writing a caller-owned stream, which remains open.</summary>
    /// <remarks>Seekable destinations are replaced and rewound after success; non-seekable destinations receive bytes at their current position. Preparation, policy and cancellation failures before publication preserve bytes and position. Stream I/O failures during publication may leave partial output. Synchronous stream cancellation is observed before publication.</remarks>
    public PublisherSaveReport Save(Stream stream, PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) {
        RequireNativeDestination(stream);
        var timer = Stopwatch.StartNew();
        byte[] bytes = PrepareSave((options ?? new PublisherSaveOptions()).Clone(), timer, cancellationToken, out var prepared);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeStreamWriter.WriteAllBytes(stream, bytes);
        return new PublisherSaveReport(SourceReadReport, prepared, timer.Elapsed, itemsWritten: 1);
    }

    /// <summary>Asynchronously stages and atomically saves to the associated native path, replacing that file by default. Explicit settings retain their selected conflict policy.</summary>
    /// <remarks>The destination must come from a successful path load or save. Native source and loss-policy requirements remain unchanged.</remarks>
    public Task<PublisherSaveReport> SaveAsync(PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        SaveAsync(FilePath ?? throw new InvalidOperationException("The publication has no associated native path. Use SaveAsync(string) or SaveAsync(Stream)."),
            options ?? new PublisherSaveOptions { FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace }, cancellationToken);

    /// <summary>Asynchronously stages and atomically publishes a complete native PUB to an explicit .pub path.</summary>
    /// <remarks>Preparation is performed in process before asynchronous I/O. Existing destinations and the associated path survive pre-publication failures. Success associates the destination without changing the original source baseline.</remarks>
    public async Task<PublisherSaveReport> SaveAsync(string path, PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) {
        string target = NativeSavePath(path);
        var timer = Stopwatch.StartNew();
        PublisherSaveOptions settings = (options ?? new PublisherSaveOptions()).Clone();
        byte[] bytes = PrepareSave(settings, timer, cancellationToken, out var prepared);
        await OfficeFileCommit.WriteAtomicallyAsync(target, (destination, token) => destination.WriteAsync(bytes, 0, bytes.Length, token),
            cancellationToken, FileConflict(settings)).ConfigureAwait(false);
        FilePath = target;
        return new PublisherSaveReport(SourceReadReport, prepared, timer.Elapsed, itemsWritten: 1, outputPath: target);
    }

    /// <summary>Asynchronously writes a complete native PUB to a caller-owned stream without closing it.</summary>
    /// <remarks>Seekable destinations are replaced and rewound after success; non-seekable destinations receive bytes at their current position. Preparation failures leave bytes and position unchanged. Cancellation or I/O failure during stream publication can leave partial output.</remarks>
    public async Task<PublisherSaveReport> SaveAsync(Stream stream, PublisherSaveOptions? options = null, CancellationToken cancellationToken = default) {
        RequireNativeDestination(stream);
        var timer = Stopwatch.StartNew();
        byte[] bytes = PrepareSave((options ?? new PublisherSaveOptions()).Clone(), timer, cancellationToken, out var prepared);
        await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
        return new PublisherSaveReport(SourceReadReport, prepared, timer.Elapsed, itemsWritten: 1);
    }

    private byte[] PrepareSave(PublisherSaveOptions settings, Stopwatch timer, CancellationToken token, out PublisherNativePreparation prepared) {
        token.ThrowIfCancellationRequested();
        prepared = PrepareNativeBytes(settings.MaxOutputBytes, token);
        if (!prepared.IsPrepared || prepared.HasErrors)
            throw new OfficeConversionException("The current publication cannot be saved through the supported native source profiles. Inspect the save report.",
                new PublisherSaveReport(SourceReadReport, prepared, timer.Elapsed));
        if (!settings.AllowLoss && prepared.HasLoss)
            new PublisherSaveReport(SourceReadReport, prepared, timer.Elapsed).ThrowIfLoss();
        byte[] bytes = prepared.CopyBytes();
        token.ThrowIfCancellationRequested();
        return bytes;
    }

    private static string NativeSavePath(string path) {
        ArgumentNullException.ThrowIfNull(path);
        if (string.IsNullOrWhiteSpace(path) || !string.Equals(Path.GetExtension(path), ".pub", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Native Publisher saving requires a .pub destination.", nameof(path));
        return Path.GetFullPath(path);
    }

    private static void RequireNativeDestination(Stream stream) {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanWrite) throw new ArgumentException("The native PUB destination must be writable.", nameof(stream));
    }

    private static OfficeFileCommit.ConflictPolicy FileConflict(PublisherSaveOptions settings) =>
        settings.FileConflictPolicy == OfficeConversionFileConflictPolicy.Replace
            ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists;
}
