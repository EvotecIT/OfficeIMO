namespace OfficeIMO.Pdf;

/// <summary>
/// Immutable identity and size evidence for one PDF artifact in a processing pipeline.
/// </summary>
public sealed class PdfArtifactSnapshot {
    private PdfArtifactSnapshot(long byteCount, string sha256, int? pageCount) {
        ByteCount = byteCount;
        Sha256 = sha256;
        PageCount = pageCount;
    }

    /// <summary>Artifact length in bytes.</summary>
    public long ByteCount { get; }

    /// <summary>Lowercase SHA-256 digest of the exact artifact bytes.</summary>
    public string Sha256 { get; }

    /// <summary>Readable page count, or null when page inspection did not complete.</summary>
    public int? PageCount { get; }

    internal static PdfArtifactSnapshot Capture(byte[] bytes, PdfReadOptions? readOptions = null) {
        Guard.NotNull(bytes, nameof(bytes));

        int? pageCount = null;
        try {
            pageCount = PdfInspector.Inspect(bytes, readOptions).PageCount;
        } catch {
            // Artifact identity remains useful even when a failed pipeline step produced unreadable bytes.
        }

        return CaptureKnownPageCount(bytes, pageCount);
    }

    /// <summary>Captures artifact identity when the owning operation already has canonical page-count evidence.</summary>
    internal static PdfArtifactSnapshot CaptureKnownPageCount(byte[] bytes, int? pageCount) {
        Guard.NotNull(bytes, nameof(bytes));
        return new PdfArtifactSnapshot(bytes.LongLength, ComputeSha256Hex(bytes), pageCount);
    }

    internal static PdfArtifactSnapshot FromDigest(long byteCount, string sha256, int? pageCount) {
#if NET8_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfNegative(byteCount);
#else
        if (byteCount < 0) {
            throw new ArgumentOutOfRangeException(nameof(byteCount));
        }
#endif

        Guard.NotNullOrWhiteSpace(sha256, nameof(sha256));
        return new PdfArtifactSnapshot(byteCount, sha256, pageCount);
    }

    private static string ComputeSha256Hex(byte[] bytes) {
#if NET8_0_OR_GREATER
        return ToLowerHex(System.Security.Cryptography.SHA256.HashData(bytes));
#else
        using (System.Security.Cryptography.SHA256 sha256 = System.Security.Cryptography.SHA256.Create()) {
            return ToLowerHex(sha256.ComputeHash(bytes));
        }
#endif
    }

    private static string ToLowerHex(byte[] bytes) {
        const string hex = "0123456789abcdef";
        char[] chars = new char[bytes.Length * 2];
        for (int i = 0; i < bytes.Length; i++) {
            chars[i * 2] = hex[bytes[i] >> 4];
            chars[(i * 2) + 1] = hex[bytes[i] & 0x0F];
        }

        return new string(chars);
    }
}
