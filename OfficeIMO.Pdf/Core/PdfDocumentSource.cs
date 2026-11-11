using OfficeIMO.Core.Internal;
using System.Runtime.ExceptionServices;

namespace OfficeIMO.Pdf;

/// <summary>
/// Owns the immutable bytes and read contract for an opened PDF.
/// Caller-owned buffers are copied once; buffers produced inside OfficeIMO.Pdf are adopted.
/// </summary>
internal sealed class PdfDocumentSource {
    private readonly byte[] _bytes;
    private readonly PdfReadCache<PdfReadDocument> _readCache = new();
    private readonly bool _canonicalRewrittenOutput;
    private ExceptionDispatchInfo? _readFailure;
    private StructuralMarkerSnapshot? _structuralMarkers;

    private PdfDocumentSource(byte[] bytes, PdfLoadOptions options, bool canonicalRewrittenOutput = false) {
        _bytes = bytes;
        Options = options.Clone();
        _canonicalRewrittenOutput = canonicalRewrittenOutput;
    }

    private PdfDocumentSource(byte[] bytes, PdfLoadOptions options, PdfReadDocument readDocument) {
        _bytes = bytes;
        Options = options.Clone();
        _readCache = new PdfReadCache<PdfReadDocument>(readDocument);
    }

    /// <summary>Immutable read settings captured when the source is opened.</summary>
    internal PdfLoadOptions Options { get; }

    /// <summary>Returns the owned source bytes for in-assembly operations without another allocation.</summary>
    internal byte[] Bytes => _bytes;

    /// <summary>
    /// Returns structural marker counts for the immutable source bytes, reusing the last scan made under the same limits.
    /// Every output derived from this source sizes its read limits from these counts.
    /// </summary>
    internal PdfStructuralMarkerCounts GetStructuralMarkers(PdfReadLimits limits) {
        Guard.NotNull(limits, nameof(limits));
        StructuralMarkerSnapshot? snapshot = Volatile.Read(ref _structuralMarkers);
        if (snapshot is not null && ReferenceEquals(snapshot.Limits, limits)) {
            return snapshot.Counts;
        }

        PdfStructuralMarkerCounts counts = PdfSyntax.InspectStructuralMarkers(_bytes, limits);
        Volatile.Write(ref _structuralMarkers, new StructuralMarkerSnapshot(limits, counts));
        return counts;
    }

    /// <summary>Copies the source bytes for a caller-owned result.</summary>
    internal byte[] CopyBytes() => (byte[])_bytes.Clone();

    /// <summary>Snapshots caller-owned bytes after enforcing the configured input budget.</summary>
    internal static PdfDocumentSource FromCallerBytes(byte[] bytes, PdfLoadOptions? options) {
        Guard.NotNull(bytes, nameof(bytes));
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        ValidateLength(bytes.LongLength, effectiveOptions);
        return new PdfDocumentSource((byte[])bytes.Clone(), effectiveOptions);
    }

    /// <summary>Adopts an internal operation result without copying it again.</summary>
    internal static PdfDocumentSource FromOwnedBytes(byte[] bytes, PdfLoadOptions? options) {
        Guard.NotNull(bytes, nameof(bytes));
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        ValidateLength(bytes.LongLength, effectiveOptions);
        return new PdfDocumentSource(bytes, effectiveOptions);
    }

    /// <summary>Adopts canonical clear-text rewrite bytes whose first semantic read can skip arbitrary-input marker discovery.</summary>
    internal static PdfDocumentSource FromOwnedRewrittenBytes(byte[] bytes, PdfLoadOptions? options) {
        Guard.NotNull(bytes, nameof(bytes));
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        ValidateLength(bytes.LongLength, effectiveOptions);
        return new PdfDocumentSource(bytes, effectiveOptions, canonicalRewrittenOutput: true);
    }

    /// <summary>Adopts internal bytes together with the canonical parse that already validated them.</summary>
    internal static PdfDocumentSource FromOwnedBytes(
        byte[] bytes,
        PdfLoadOptions? options,
        PdfReadDocument readDocument) {
        Guard.NotNull(bytes, nameof(bytes));
        Guard.NotNull(readDocument, nameof(readDocument));
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        ValidateLength(bytes.LongLength, effectiveOptions);
        return new PdfDocumentSource(bytes, effectiveOptions, readDocument);
    }

    /// <summary>Reads and owns one bounded file snapshot.</summary>
    internal static PdfDocumentSource FromPath(string path, PdfLoadOptions? options) =>
        FromPath(path, options, CancellationToken.None);

    /// <summary>Reads and owns one bounded file snapshot with cooperative cancellation.</summary>
    internal static PdfDocumentSource FromPath(
        string path,
        PdfLoadOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        string fullPath = Path.GetFullPath(path);
        var file = new FileInfo(fullPath);
        ValidateLength(file.Length, effectiveOptions);

        using var stream = new FileStream(
            fullPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        return FromBoundedStream(stream, effectiveOptions, cancellationToken);
    }

    /// <summary>
    /// Reads and owns one bounded stream snapshot. Seekable streams are read from the beginning and restored.
    /// </summary>
    internal static PdfDocumentSource FromStream(Stream stream, PdfLoadOptions? options, CancellationToken cancellationToken = default) {
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        return FromBoundedStream(stream, effectiveOptions, cancellationToken);
    }

    /// <summary>
    /// Reads and owns one bounded stream snapshot from the caller's current position.
    /// </summary>
    internal static PdfDocumentSource FromRemainingStream(Stream stream, PdfLoadOptions? options) =>
        FromRemainingStream(stream, options, CancellationToken.None);

    /// <summary>
    /// Reads and owns one bounded stream snapshot from the caller's current position with cooperative cancellation.
    /// </summary>
    internal static PdfDocumentSource FromRemainingStream(
        Stream stream,
        PdfLoadOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        long limit = effectiveOptions.Limits.MaxInputBytes;
        try {
            byte[] bytes = OfficeStreamReader.ReadRemainingBytes(stream, cancellationToken, limit);
            return FromOwnedBytes(bytes, effectiveOptions);
        } catch (InvalidDataException exception) when (OfficeStreamReader.IsSizeLimitException(exception)) {
            throw CreateInputLimitException(stream, limit, remainingOnly: true);
        }
    }

    /// <summary>Asynchronously reads and owns one bounded file snapshot.</summary>
    internal static async Task<PdfDocumentSource> FromPathAsync(
        string path,
        PdfLoadOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        string fullPath = Path.GetFullPath(path);
        var file = new FileInfo(fullPath);
        ValidateLength(file.Length, effectiveOptions);

        using var stream = new FileStream(
            fullPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete,
            81920,
            useAsync: true);
        return await FromBoundedStreamAsync(stream, effectiveOptions, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Asynchronously reads and owns one bounded stream snapshot. Seekable streams are read from the beginning and restored.
    /// </summary>
    internal static Task<PdfDocumentSource> FromStreamAsync(
        Stream stream,
        PdfLoadOptions? options,
        CancellationToken cancellationToken) {
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        return FromBoundedStreamAsync(stream, effectiveOptions, cancellationToken);
    }

    /// <summary>Returns the cached canonical parse or a one-off parse for explicit override settings.</summary>
    internal PdfReadDocument Read(
        PdfLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (options is not null && !ReferenceEquals(options, Options)) {
            return _canonicalRewrittenOutput
                ? PdfReadDocument.OpenRewrittenOutput(_bytes, options, cancellationToken)
                : PdfReadDocument.OpenOwned(_bytes, options, cancellationToken);
        }

        return _readCache.GetOrCreate(this, static (source, token) => source.ParseCanonical(token), cancellationToken);
    }

    private PdfReadDocument ParseCanonical(CancellationToken cancellationToken) {
        _readFailure?.Throw();
        try {
            return _canonicalRewrittenOutput
                ? PdfReadDocument.OpenRewrittenOutput(_bytes, Options, cancellationToken)
                : PdfReadDocument.OpenOwned(_bytes, Options, cancellationToken);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            // A caller's cancellation must not poison the reusable source cache.
            throw;
        } catch (Exception exception) {
            _readFailure = ExceptionDispatchInfo.Capture(exception);
            throw;
        }
    }

    /// <summary>
    /// Captures the opened artifact while priming and reusing the source's canonical parse.
    /// Invalid input still produces size and on-demand hash evidence and caches the parse failure.
    /// </summary>
    internal PdfArtifactInfo CaptureArtifact(CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        int? pageCount = null;
        try {
            pageCount = Read(cancellationToken: cancellationToken).Pages.Count;
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch {
            // Artifact identity remains useful even when the canonical parse fails.
        }

        return PdfArtifactInfo.CaptureOwnedKnownPageCount(_bytes, pageCount);
    }

    private static PdfDocumentSource FromBoundedStream(Stream stream, PdfLoadOptions options) =>
        FromBoundedStream(stream, options, CancellationToken.None);

    private static PdfDocumentSource FromBoundedStream(
        Stream stream,
        PdfLoadOptions options,
        CancellationToken cancellationToken) {
        Guard.NotNull(stream, nameof(stream));
        long limit = options.Limits.MaxInputBytes;
        try {
            byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, limit);
            return FromOwnedBytes(bytes, options);
        } catch (InvalidDataException exception) when (OfficeStreamReader.IsSizeLimitException(exception)) {
            throw CreateInputLimitException(stream, limit);
        }
    }

    private static async Task<PdfDocumentSource> FromBoundedStreamAsync(
        Stream stream,
        PdfLoadOptions options,
        CancellationToken cancellationToken) {
        Guard.NotNull(stream, nameof(stream));
        long limit = options.Limits.MaxInputBytes;
        try {
            byte[] bytes = await OfficeStreamReader
                .ReadAllBytesAsync(stream, cancellationToken, limit)
                .ConfigureAwait(false);
            return FromOwnedBytes(bytes, options);
        } catch (InvalidDataException exception) when (OfficeStreamReader.IsSizeLimitException(exception)) {
            throw CreateInputLimitException(stream, limit);
        }
    }

    private static void ValidateLength(long length, PdfLoadOptions options) {
        options.Limits.Validate();
        if (length > options.Limits.MaxInputBytes) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.InputBytes,
                options.Limits.MaxInputBytes,
                length);
        }
    }

    private static PdfReadLimitException CreateInputLimitException(Stream stream, long limit, bool remainingOnly = false) {
        long actual = limit + 1;
        if (stream.CanSeek) {
            try {
                actual = Math.Max(actual, remainingOnly ? stream.Length - stream.Position : stream.Length);
            } catch (NotSupportedException) {
                // The bounded reader already proved the limit was exceeded.
            }
        }

        return PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limit, actual);
    }

    private sealed class StructuralMarkerSnapshot {
        internal StructuralMarkerSnapshot(PdfReadLimits limits, PdfStructuralMarkerCounts counts) {
            Limits = limits;
            Counts = counts;
        }

        internal PdfReadLimits Limits { get; }

        internal PdfStructuralMarkerCounts Counts { get; }
    }
}
