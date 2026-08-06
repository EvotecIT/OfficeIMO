using System;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>
/// Contains image content retrieved from an HTTP or HTTPS source.
/// </summary>
public sealed class OfficeRemoteImage {
    private readonly byte[] _bytes;

    internal OfficeRemoteImage(Uri source, byte[] bytes, string fileName, string contentType) {
        Source = source;
        _bytes = bytes ?? throw new ArgumentNullException(nameof(bytes));
        FileName = fileName ?? throw new ArgumentNullException(nameof(fileName));
        ContentType = contentType ?? throw new ArgumentNullException(nameof(contentType));
    }

    /// <summary>The final source URI after any accepted redirects.</summary>
    public Uri Source { get; }

    /// <summary>A suggested file name derived from the source URI.</summary>
    public string FileName { get; }

    /// <summary>The normalized image MIME content type returned by the server.</summary>
    public string ContentType { get; }

    /// <summary>Returns a defensive copy of the downloaded image bytes.</summary>
    public byte[] ToBytes() => (byte[])_bytes.Clone();

    /// <summary>Returns a new readable stream containing the downloaded image.</summary>
    public MemoryStream ToStream() => new MemoryStream(_bytes, writable: false);
}
