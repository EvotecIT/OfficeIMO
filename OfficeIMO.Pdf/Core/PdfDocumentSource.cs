using OfficeIMO.Core.Internal;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

/// <summary>
/// Owns the immutable bytes and read contract for an opened PDF.
/// Caller-owned buffers are copied once; buffers produced inside OfficeIMO.Pdf are adopted.
/// </summary>
internal sealed class PdfDocumentSource {
    private readonly byte[] _bytes;
    private readonly object _readLock = new object();
    private PdfReadDocument? _readDocument;
    private ExceptionDispatchInfo? _readFailure;

    private PdfDocumentSource(byte[] bytes, PdfLoadOptions options) {
        _bytes = bytes;
        Options = options;
    }

    private PdfDocumentSource(byte[] bytes, PdfLoadOptions options, PdfReadDocument readDocument) {
        _bytes = bytes;
        Options = options;
        _readDocument = readDocument;
    }

    /// <summary>Immutable read settings captured when the source is opened.</summary>
    internal PdfLoadOptions Options { get; }

    /// <summary>Returns the owned source bytes for in-assembly operations without another allocation.</summary>
    internal byte[] Bytes => _bytes;

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
    internal static PdfDocumentSource FromPath(string path, PdfLoadOptions? options) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        string fullPath = Path.GetFullPath(path);
        var file = new FileInfo(fullPath);
        ValidateLength(file.Length, effectiveOptions);

        using var stream = new FileStream(
            fullPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        return FromBoundedStream(stream, effectiveOptions);
    }

    /// <summary>
    /// Reads and owns one bounded stream snapshot. Seekable streams are read from the beginning and restored.
    /// </summary>
    internal static PdfDocumentSource FromStream(Stream stream, PdfLoadOptions? options) {
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        return FromBoundedStream(stream, effectiveOptions);
    }

    /// <summary>
    /// Reads and owns one bounded stream snapshot from the caller's current position.
    /// </summary>
    internal static PdfDocumentSource FromRemainingStream(Stream stream, PdfLoadOptions? options) {
        Guard.NotNull(stream, nameof(stream));
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        long limit = effectiveOptions.Limits.MaxInputBytes;
        try {
            byte[] bytes = OfficeStreamReader.ReadRemainingBytes(stream, limit);
            return FromOwnedBytes(bytes, effectiveOptions);
        } catch (InvalidDataException) {
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
            return PdfReadDocument.Open(_bytes, options, cancellationToken);
        }

        lock (_readLock) {
            if (_readDocument is not null) return _readDocument;
            _readFailure?.Throw();
            try {
                _readDocument = PdfReadDocument.Open(_bytes, Options, cancellationToken);
                return _readDocument;
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                // A caller's cancellation must not poison the reusable source cache.
                throw;
            } catch (Exception exception) {
                _readFailure = ExceptionDispatchInfo.Capture(exception);
                throw;
            }
        }
    }

    /// <summary>
    /// Captures the opened artifact while priming and reusing the source's canonical parse.
    /// Invalid input still produces hash and size evidence and caches the parse failure.
    /// </summary>
    internal PdfArtifactSnapshot CaptureArtifact() {
        int? pageCount = null;
        try {
            pageCount = Read().Pages.Count;
        } catch {
            // Artifact identity remains useful even when the canonical parse fails.
        }

        return PdfArtifactSnapshot.CaptureKnownPageCount(_bytes, pageCount);
    }

    private static PdfDocumentSource FromBoundedStream(Stream stream, PdfLoadOptions options) {
        Guard.NotNull(stream, nameof(stream));
        long limit = options.Limits.MaxInputBytes;
        try {
            byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, limit);
            return FromOwnedBytes(bytes, options);
        } catch (InvalidDataException) {
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
        } catch (InvalidDataException) {
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
                actual = remainingOnly ? stream.Length - stream.Position : stream.Length;
            } catch (NotSupportedException) {
                // The bounded reader already proved the limit was exceeded.
            }
        }

        return PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limit, actual);
    }
}
