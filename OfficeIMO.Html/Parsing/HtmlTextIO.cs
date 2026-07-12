using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Html;

/// <summary>
/// Consistent UTF-8 text I/O for OfficeIMO HTML adapters. The default encoding is UTF-8 without a byte-order mark.
/// </summary>
public static class HtmlTextIO {
    private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

    /// <summary>Reads HTML text while detecting a byte-order mark and leaving the source stream open.</summary>
    public static string Read(Stream stream) {
        ValidateReadable(stream);
        using var reader = new StreamReader(stream, Utf8NoBom, true, 4096, true);
        return reader.ReadToEnd();
    }

    /// <summary>Asynchronously reads HTML text while leaving the source stream open.</summary>
    public static async Task<string> ReadAsync(Stream stream, CancellationToken cancellationToken = default) {
        ValidateReadable(stream);
        cancellationToken.ThrowIfCancellationRequested();
        using var reader = new StreamReader(stream, Utf8NoBom, true, 4096, true);
#if NET8_0_OR_GREATER
        return await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
#else
        string text = await reader.ReadToEndAsync().ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return text;
#endif
    }

    /// <summary>Writes HTML text to a file as UTF-8 without a byte-order mark.</summary>
    public static void Write(string path, string html) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An HTML output path is required.", nameof(path));
        OfficeFileCommit.WriteAllBytes(path, Utf8NoBom.GetBytes(html ?? throw new ArgumentNullException(nameof(html))));
    }

    /// <summary>Writes HTML text to a stream as UTF-8 without a byte-order mark and leaves it open.</summary>
    public static void Write(Stream stream, string html) {
        ValidateWritable(stream);
        byte[] bytes = Utf8NoBom.GetBytes(html ?? throw new ArgumentNullException(nameof(html)));
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>Asynchronously writes HTML text to a file as UTF-8 without a byte-order mark.</summary>
    public static async Task WriteAsync(string path, string html, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An HTML output path is required.", nameof(path));
        byte[] bytes = Utf8NoBom.GetBytes(html ?? throw new ArgumentNullException(nameof(html)));
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true);
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes HTML text to a stream as UTF-8 without a byte-order mark and leaves it open.</summary>
    public static async Task WriteAsync(Stream stream, string html, CancellationToken cancellationToken = default) {
        ValidateWritable(stream);
        byte[] bytes = Utf8NoBom.GetBytes(html ?? throw new ArgumentNullException(nameof(html)));
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
    }

    private static void ValidateReadable(Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The HTML input stream must be readable.", nameof(stream));
    }

    private static void ValidateWritable(Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The HTML output stream must be writable.", nameof(stream));
    }
}
