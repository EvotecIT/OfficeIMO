using System;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Original Office source retained alongside a visual or editable compatibility fallback.</summary>
public sealed class OfficeCompatibilitySourcePayload {
    private readonly byte[] _bytes;

    internal OfficeCompatibilitySourcePayload(
        string formatId,
        string fileName,
        string sha256,
        OfficeCompatibilityMode mode,
        byte[] bytes) {
        FormatId = formatId ?? throw new ArgumentNullException(nameof(formatId));
        FileName = fileName ?? throw new ArgumentNullException(nameof(fileName));
        Sha256 = sha256 ?? throw new ArgumentNullException(nameof(sha256));
        Mode = mode;
        _bytes = bytes ?? throw new ArgumentNullException(nameof(bytes));
    }

    /// <summary>Gets the concrete OfficeIMO format identifier of the retained source.</summary>
    public string FormatId { get; }

    /// <summary>Gets the source file name recorded at conversion time.</summary>
    public string FileName { get; }

    /// <summary>Gets the lower-case SHA-256 digest of the retained bytes.</summary>
    public string Sha256 { get; }

    /// <summary>Gets the compatibility mode that created the carrier.</summary>
    public OfficeCompatibilityMode Mode { get; }

    /// <summary>Gets the retained source length.</summary>
    public long Length => _bytes.LongLength;

    /// <summary>Opens an independent read-only stream over the retained source.</summary>
    public Stream OpenRead() => new MemoryStream(_bytes, writable: false);

    /// <summary>Returns an independent copy of the retained source bytes.</summary>
    public byte[] ToArray() => (byte[])_bytes.Clone();
}
