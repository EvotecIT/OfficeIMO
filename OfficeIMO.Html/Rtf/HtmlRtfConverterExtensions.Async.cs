namespace OfficeIMO.Html;

/// <content>
/// Provides asynchronous IO extension methods for converting between RTF and semantic HTML.
/// </content>
public static partial class HtmlRtfConverterExtensions {
    /// <summary>Converts an RTF document to semantic HTML.</summary>
    public static Task<string> ToHtmlAsync(this RtfDocument document, RtfToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(document.ToHtml(options));
    }

    /// <summary>Converts an RTF document to encoded semantic HTML bytes.</summary>
    public static Task<byte[]> ToHtmlBytesAsync(this RtfDocument document, RtfToHtmlOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(document.ToHtmlBytes(options, encoding));
    }

    /// <summary>Converts an RTF document to an encoded semantic HTML memory stream.</summary>
    public static Task<MemoryStream> ToHtmlMemoryStreamAsync(this RtfDocument document, RtfToHtmlOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(document.ToHtmlMemoryStream(options, encoding));
    }

    /// <summary>Saves an RTF document as semantic HTML at the specified path.</summary>
    public static async Task SaveAsHtmlAsync(this RtfDocument document, string path, RtfToHtmlOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        cancellationToken.ThrowIfCancellationRequested();
        await WriteTextAsync(path, document.ToHtml(options), encoding ?? Encoding.UTF8, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves an RTF document as semantic HTML to a writable stream.</summary>
    public static async Task SaveAsHtmlAsync(this RtfDocument document, Stream stream, RtfToHtmlOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        cancellationToken.ThrowIfCancellationRequested();
        await WriteBytesAsync(stream, document.ToHtmlBytes(options, encoding), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Loads semantic HTML into an RTF document model.</summary>
    public static Task<RtfDocument> LoadFromHtmlAsync(this string html, HtmlToRtfOptions? options = null, CancellationToken cancellationToken = default) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(html.LoadFromHtml(options));
    }

    /// <summary>Loads encoded semantic HTML bytes into an RTF document model.</summary>
    public static Task<RtfDocument> LoadFromHtmlAsync(this byte[] htmlBytes, HtmlToRtfOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (htmlBytes == null) {
            throw new ArgumentNullException(nameof(htmlBytes));
        }

        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(htmlBytes.LoadFromHtml(options, encoding));
    }

    /// <summary>Reads semantic HTML from a stream into an RTF document model.</summary>
    public static async Task<RtfDocument> LoadFromHtmlAsync(this Stream htmlStream, HtmlToRtfOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        string html = await ReadTextAsync(htmlStream, encoding ?? Encoding.UTF8, cancellationToken).ConfigureAwait(false);
        return html.LoadFromHtml(options);
    }

    /// <summary>Converts semantic HTML to RTF text.</summary>
    public static async Task<string> ToRtfAsync(this string html, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, CancellationToken cancellationToken = default) {
        RtfDocument document = await html.LoadFromHtmlAsync(readOptions, cancellationToken).ConfigureAwait(false);
        return document.ToRtf(writeOptions);
    }

    /// <summary>Converts encoded semantic HTML bytes to RTF text.</summary>
    public static async Task<string> ToRtfAsync(this byte[] htmlBytes, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        RtfDocument document = await htmlBytes.LoadFromHtmlAsync(readOptions, encoding, cancellationToken).ConfigureAwait(false);
        return document.ToRtf(writeOptions);
    }

    /// <summary>Reads semantic HTML from a stream and converts it to RTF text.</summary>
    public static async Task<string> ToRtfAsync(this Stream htmlStream, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        RtfDocument document = await htmlStream.LoadFromHtmlAsync(readOptions, encoding, cancellationToken).ConfigureAwait(false);
        return document.ToRtf(writeOptions);
    }

    /// <summary>Converts semantic HTML to encoded RTF bytes.</summary>
    public static async Task<byte[]> ToRtfBytesAsync(this string html, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        string rtf = await html.ToRtfAsync(readOptions, writeOptions, cancellationToken).ConfigureAwait(false);
        return (encoding ?? Encoding.UTF8).GetBytes(rtf);
    }

    /// <summary>Converts encoded semantic HTML bytes to encoded RTF bytes.</summary>
    public static async Task<byte[]> ToRtfBytesAsync(this byte[] htmlBytes, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null, CancellationToken cancellationToken = default) {
        string rtf = await htmlBytes.ToRtfAsync(readOptions, writeOptions, htmlEncoding, cancellationToken).ConfigureAwait(false);
        return (rtfEncoding ?? Encoding.UTF8).GetBytes(rtf);
    }

    /// <summary>Reads semantic HTML from a stream and converts it to encoded RTF bytes.</summary>
    public static async Task<byte[]> ToRtfBytesAsync(this Stream htmlStream, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null, CancellationToken cancellationToken = default) {
        string rtf = await htmlStream.ToRtfAsync(readOptions, writeOptions, htmlEncoding, cancellationToken).ConfigureAwait(false);
        return (rtfEncoding ?? Encoding.UTF8).GetBytes(rtf);
    }

    /// <summary>Converts semantic HTML to an encoded RTF memory stream.</summary>
    public static async Task<MemoryStream> ToRtfMemoryStreamAsync(this string html, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        byte[] bytes = await html.ToRtfBytesAsync(readOptions, writeOptions, encoding, cancellationToken).ConfigureAwait(false);
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>Converts encoded semantic HTML bytes to an encoded RTF memory stream.</summary>
    public static async Task<MemoryStream> ToRtfMemoryStreamAsync(this byte[] htmlBytes, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null, CancellationToken cancellationToken = default) {
        byte[] bytes = await htmlBytes.ToRtfBytesAsync(readOptions, writeOptions, htmlEncoding, rtfEncoding, cancellationToken).ConfigureAwait(false);
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>Reads semantic HTML from a stream and converts it to an encoded RTF memory stream.</summary>
    public static async Task<MemoryStream> ToRtfMemoryStreamAsync(this Stream htmlStream, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null, CancellationToken cancellationToken = default) {
        byte[] bytes = await htmlStream.ToRtfBytesAsync(readOptions, writeOptions, htmlEncoding, rtfEncoding, cancellationToken).ConfigureAwait(false);
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>Saves semantic HTML as an RTF file at the specified path.</summary>
    public static async Task SaveAsRtfAsync(this string html, string path, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        byte[] bytes = await html.ToRtfBytesAsync(readOptions, writeOptions, encoding, cancellationToken).ConfigureAwait(false);
        await WriteBytesAsync(path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves encoded semantic HTML bytes as an RTF file at the specified path.</summary>
    public static async Task SaveAsRtfAsync(this byte[] htmlBytes, string path, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null, CancellationToken cancellationToken = default) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        byte[] bytes = await htmlBytes.ToRtfBytesAsync(readOptions, writeOptions, htmlEncoding, rtfEncoding, cancellationToken).ConfigureAwait(false);
        await WriteBytesAsync(path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reads semantic HTML from a stream and saves it as an RTF file at the specified path.</summary>
    public static async Task SaveAsRtfAsync(this Stream htmlStream, string path, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null, CancellationToken cancellationToken = default) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        byte[] bytes = await htmlStream.ToRtfBytesAsync(readOptions, writeOptions, htmlEncoding, rtfEncoding, cancellationToken).ConfigureAwait(false);
        await WriteBytesAsync(path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves semantic HTML as RTF to a writable stream without closing or rewinding the stream.</summary>
    public static async Task SaveAsRtfAsync(this string html, Stream stream, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        byte[] bytes = await html.ToRtfBytesAsync(readOptions, writeOptions, encoding, cancellationToken).ConfigureAwait(false);
        await WriteBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves encoded semantic HTML bytes as RTF to a writable stream without closing or rewinding the stream.</summary>
    public static async Task SaveAsRtfAsync(this byte[] htmlBytes, Stream stream, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null, CancellationToken cancellationToken = default) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        byte[] bytes = await htmlBytes.ToRtfBytesAsync(readOptions, writeOptions, htmlEncoding, rtfEncoding, cancellationToken).ConfigureAwait(false);
        await WriteBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reads semantic HTML from a stream and saves it as RTF to a writable stream without closing or rewinding the stream.</summary>
    public static async Task SaveAsRtfAsync(this Stream htmlStream, Stream stream, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null, CancellationToken cancellationToken = default) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        byte[] bytes = await htmlStream.ToRtfBytesAsync(readOptions, writeOptions, htmlEncoding, rtfEncoding, cancellationToken).ConfigureAwait(false);
        await WriteBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    private static async Task<string> ReadTextAsync(Stream stream, Encoding encoding, CancellationToken cancellationToken) {
        using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
#if NET8_0_OR_GREATER
        return await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
#else
        cancellationToken.ThrowIfCancellationRequested();
        return await reader.ReadToEndAsync().ConfigureAwait(false);
#endif
    }

    private static async Task WriteTextAsync(string path, string text, Encoding encoding, CancellationToken cancellationToken) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

#if NET8_0_OR_GREATER
        await File.WriteAllTextAsync(path, text, encoding, cancellationToken).ConfigureAwait(false);
#else
        cancellationToken.ThrowIfCancellationRequested();
        using var writer = new StreamWriter(path, append: false, encoding);
        await writer.WriteAsync(text).ConfigureAwait(false);
#endif
    }

    private static async Task WriteBytesAsync(string path, byte[] bytes, CancellationToken cancellationToken) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

#if NET8_0_OR_GREATER
        await File.WriteAllBytesAsync(path, bytes, cancellationToken).ConfigureAwait(false);
#else
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
    }

    private static async Task WriteBytesAsync(Stream stream, byte[] bytes, CancellationToken cancellationToken) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

#if NET8_0_OR_GREATER
        await stream.WriteAsync(bytes, cancellationToken).ConfigureAwait(false);
#else
        cancellationToken.ThrowIfCancellationRequested();
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
    }
}
