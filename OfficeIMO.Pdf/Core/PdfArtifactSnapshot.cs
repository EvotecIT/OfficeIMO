namespace OfficeIMO.Pdf;

/// <summary>
/// Immutable identity and size evidence for one PDF artifact in a processing pipeline.
/// </summary>
public sealed class PdfArtifactInfo {
    // Holds the digest string, or the immutable owned bytes until the digest is first requested.
    private object _digestState;

    private PdfArtifactInfo(long byteCount, string sha256, int? pageCount) {
        ByteCount = byteCount;
        _digestState = sha256;
        PageCount = pageCount;
    }

    private PdfArtifactInfo(ReadOnlyMemory<byte> ownedBytes, int? pageCount) {
        ByteCount = ownedBytes.Length;
        _digestState = ownedBytes;
        PageCount = pageCount;
    }

    /// <summary>Artifact length in bytes.</summary>
    public long ByteCount { get; }

    /// <summary>
    /// Lowercase SHA-256 digest of the exact artifact bytes. Artifacts captured from bytes owned by a
    /// <see cref="PdfDocument"/> compute the digest on first access and then release the retained bytes.
    /// </summary>
    public string Sha256 {
        get {
            object state = Volatile.Read(ref _digestState);
            if (state is string digest) {
                return digest;
            }

            string computed = ComputeSha256Hex((ReadOnlyMemory<byte>)state, CancellationToken.None);
            Interlocked.CompareExchange(ref _digestState, computed, state);
            return (string)Volatile.Read(ref _digestState);
        }
    }

    /// <summary>Readable page count, or null when page inspection did not complete.</summary>
    public int? PageCount { get; }

    internal static PdfArtifactInfo Capture(byte[] bytes, PdfLoadOptions? readOptions = null,
        CancellationToken cancellationToken = default) =>
        Capture(bytes, readOptions, out _, cancellationToken);

    /// <summary>Retains a successful artifact readback for the next operation on the same output.</summary>
    internal static PdfArtifactInfo Capture(
        byte[] bytes,
        PdfLoadOptions? readOptions,
        out PdfReadDocument? readDocument,
        CancellationToken cancellationToken = default) {
        int? pageCount = ReadPageCount(bytes, readOptions, out readDocument, cancellationToken);
        return CaptureKnownPageCount(bytes, pageCount, cancellationToken);
    }

    /// <summary>Reads page-count evidence for bytes owned by a document and defers their digest until requested.</summary>
    internal static PdfArtifactInfo CaptureOwned(
        byte[] ownedBytes,
        PdfLoadOptions? readOptions,
        out PdfReadDocument? readDocument,
        CancellationToken cancellationToken = default) {
        int? pageCount = ReadPageCount(ownedBytes, readOptions, out readDocument, cancellationToken);
        return CaptureOwnedKnownPageCount(ownedBytes, pageCount);
    }

    private static int? ReadPageCount(
        byte[] bytes,
        PdfLoadOptions? readOptions,
        out PdfReadDocument? readDocument,
        CancellationToken cancellationToken) {
        Guard.NotNull(bytes, nameof(bytes));
        cancellationToken.ThrowIfCancellationRequested();

        int? pageCount = null;
        readDocument = null;
        try {
            PdfReadDocument parsed = PdfReadDocument.Open(bytes, readOptions, cancellationToken);
            pageCount = parsed.Pages.Count;
            readDocument = parsed;
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch {
            // Artifact identity remains useful even when a failed pipeline step produced unreadable bytes.
        }

        return pageCount;
    }

    /// <summary>Captures artifact identity when the owning operation already has canonical page-count evidence.</summary>
    internal static PdfArtifactInfo CaptureKnownPageCount(byte[] bytes, int? pageCount,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(bytes, nameof(bytes));
        return new PdfArtifactInfo(bytes.LongLength, ComputeSha256Hex(bytes, cancellationToken), pageCount);
    }

    internal static PdfArtifactInfo CaptureKnownPageCount(ReadOnlyMemory<byte> bytes, int? pageCount,
        CancellationToken cancellationToken) =>
        new PdfArtifactInfo(bytes.Length, ComputeSha256Hex(bytes, cancellationToken), pageCount);

    /// <summary>
    /// Captures identity for bytes that no code will mutate for the lifetime of the artifact, deferring the digest until requested.
    /// </summary>
    internal static PdfArtifactInfo CaptureOwnedKnownPageCount(byte[] ownedBytes, int? pageCount) {
        Guard.NotNull(ownedBytes, nameof(ownedBytes));
        return CaptureOwnedKnownPageCount((ReadOnlyMemory<byte>)ownedBytes, pageCount);
    }

    internal static PdfArtifactInfo CaptureOwnedKnownPageCount(ReadOnlyMemory<byte> ownedBytes, int? pageCount) =>
        new PdfArtifactInfo(ownedBytes, pageCount);

    internal static PdfArtifactInfo FromDigest(long byteCount, string sha256, int? pageCount) {
        ArgumentOutOfRangeException.ThrowIfNegative(byteCount);

        Guard.NotNullOrWhiteSpace(sha256, nameof(sha256));
        return new PdfArtifactInfo(byteCount, sha256, pageCount);
    }

    internal static string ComputeSha256Hex(ReadOnlyMemory<byte> bytes, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
#if NET8_0_OR_GREATER
        if (!cancellationToken.CanBeCanceled)
            return ToLowerHex(System.Security.Cryptography.SHA256.HashData(bytes.Span));
#endif
        using var sha256 = System.Security.Cryptography.IncrementalHash.CreateHash(
            System.Security.Cryptography.HashAlgorithmName.SHA256);
        const int chunkSize = 65536;
        ReadOnlySpan<byte> input = bytes.Span;
#if !NET8_0_OR_GREATER
        System.Runtime.InteropServices.MemoryMarshal.TryGetArray(bytes, out ArraySegment<byte> segment);
        byte[] buffer = segment.Array ?? new byte[Math.Min(input.Length, chunkSize)];
#endif
        for (int offset = 0; offset < input.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(chunkSize, input.Length - offset);
#if NET8_0_OR_GREATER
            sha256.AppendData(input.Slice(offset, count));
#else
            if (segment.Array is null) {
                input.Slice(offset, count).CopyTo(buffer);
                sha256.AppendData(buffer, 0, count);
            } else {
                sha256.AppendData(buffer, segment.Offset + offset, count);
            }
#endif
            offset += count;
        }

        cancellationToken.ThrowIfCancellationRequested();
        return ToLowerHex(sha256.GetHashAndReset());
    }

    internal static string ToLowerHex(byte[] bytes) {
        const string hex = "0123456789abcdef";
        char[] chars = new char[bytes.Length * 2];
        for (int i = 0; i < bytes.Length; i++) {
            chars[i * 2] = hex[bytes[i] >> 4];
            chars[(i * 2) + 1] = hex[bytes[i] & 0x0F];
        }

        return new string(chars);
    }
}
