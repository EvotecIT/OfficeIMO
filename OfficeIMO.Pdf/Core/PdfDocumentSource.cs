using OfficeIMO.Drawing.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

/// <summary>
/// Owns the immutable bytes and read contract for an opened PDF.
/// Caller-owned buffers are copied once; buffers produced inside OfficeIMO.Pdf are adopted.
/// </summary>
internal sealed class PdfDocumentSource {
    private readonly byte[] _bytes;
    private readonly Lazy<PdfReadDocument> _readDocument;

    private PdfDocumentSource(byte[] bytes, PdfReadOptions options) {
        _bytes = bytes;
        Options = options;
        _readDocument = new Lazy<PdfReadDocument>(
            () => PdfReadDocument.Open(_bytes, Options),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);
    }

    /// <summary>Immutable read settings captured when the source is opened.</summary>
    internal PdfReadOptions Options { get; }

    /// <summary>Returns the owned source bytes for in-assembly operations without another allocation.</summary>
    internal byte[] Bytes => _bytes;

    /// <summary>Copies the source bytes for a caller-owned result.</summary>
    internal byte[] CopyBytes() => (byte[])_bytes.Clone();

    /// <summary>Snapshots caller-owned bytes after enforcing the configured input budget.</summary>
    internal static PdfDocumentSource FromCallerBytes(byte[] bytes, PdfReadOptions? options) {
        Guard.NotNull(bytes, nameof(bytes));
        PdfReadOptions effectiveOptions = PdfReadOptions.Resolve(options);
        ValidateLength(bytes.LongLength, effectiveOptions);
        return new PdfDocumentSource((byte[])bytes.Clone(), effectiveOptions);
    }

    /// <summary>Adopts an internal operation result without copying it again.</summary>
    internal static PdfDocumentSource FromOwnedBytes(byte[] bytes, PdfReadOptions? options) {
        Guard.NotNull(bytes, nameof(bytes));
        PdfReadOptions effectiveOptions = PdfReadOptions.Resolve(options);
        ValidateLength(bytes.LongLength, effectiveOptions);
        return new PdfDocumentSource(bytes, effectiveOptions);
    }

    /// <summary>Reads and owns one bounded file snapshot.</summary>
    internal static PdfDocumentSource FromPath(string path, PdfReadOptions? options) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        PdfReadOptions effectiveOptions = PdfReadOptions.Resolve(options);
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
    internal static PdfDocumentSource FromStream(Stream stream, PdfReadOptions? options) {
        PdfReadOptions effectiveOptions = PdfReadOptions.Resolve(options);
        return FromBoundedStream(stream, effectiveOptions);
    }

    /// <summary>Asynchronously reads and owns one bounded file snapshot.</summary>
    internal static async Task<PdfDocumentSource> FromPathAsync(
        string path,
        PdfReadOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        PdfReadOptions effectiveOptions = PdfReadOptions.Resolve(options);
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
        PdfReadOptions? options,
        CancellationToken cancellationToken) {
        PdfReadOptions effectiveOptions = PdfReadOptions.Resolve(options);
        return FromBoundedStreamAsync(stream, effectiveOptions, cancellationToken);
    }

    /// <summary>Returns the cached canonical parse or a one-off parse for explicit override settings.</summary>
    internal PdfReadDocument Read(PdfReadOptions? options = null) {
        if (options is null || ReferenceEquals(options, Options)) {
            return _readDocument.Value;
        }

        return PdfReadDocument.Open(_bytes, options);
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

    private static PdfDocumentSource FromBoundedStream(Stream stream, PdfReadOptions options) {
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
        PdfReadOptions options,
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

    private static void ValidateLength(long length, PdfReadOptions options) {
        options.Limits.Validate();
        if (length > options.Limits.MaxInputBytes) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.InputBytes,
                options.Limits.MaxInputBytes,
                length);
        }
    }

    private static PdfReadLimitException CreateInputLimitException(Stream stream, long limit) {
        long actual = limit + 1;
        if (stream.CanSeek) {
            try {
                actual = stream.Length;
            } catch (NotSupportedException) {
                // The bounded reader already proved the limit was exceeded.
            }
        }

        return PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limit, actual);
    }
}
