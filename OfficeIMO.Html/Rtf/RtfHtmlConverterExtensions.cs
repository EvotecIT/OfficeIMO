namespace OfficeIMO.Html;

/// <summary>
/// Provides extension methods for converting between <see cref="RtfDocument"/> and semantic HTML.
/// </summary>
public static partial class RtfHtmlConverterExtensions {
    /// <summary>Converts an RTF document to semantic HTML.</summary>
    public static string ToHtml(this RtfDocument document, RtfHtmlSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        return RtfHtmlWriter.Write(document, options ?? new RtfHtmlSaveOptions());
    }

    /// <summary>Converts an RTF document to encoded semantic HTML bytes.</summary>
    public static byte[] ToHtmlBytes(this RtfDocument document, RtfHtmlSaveOptions? options = null, Encoding? encoding = null) {
        return (encoding ?? Encoding.UTF8).GetBytes(document.ToHtml(options));
    }

    /// <summary>Converts an RTF document to an encoded semantic HTML memory stream.</summary>
    public static MemoryStream ToHtmlMemoryStream(this RtfDocument document, RtfHtmlSaveOptions? options = null, Encoding? encoding = null) {
        return new MemoryStream(document.ToHtmlBytes(options, encoding), writable: false);
    }

    /// <summary>Saves an RTF document as semantic HTML at the specified path.</summary>
    public static void SaveAsHtml(this RtfDocument document, string path, RtfHtmlSaveOptions? options = null, Encoding? encoding = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        File.WriteAllText(path, document.ToHtml(options), encoding ?? Encoding.UTF8);
    }

    /// <summary>Saves an RTF document as semantic HTML to a writable stream.</summary>
    public static void SaveAsHtml(this RtfDocument document, Stream stream, RtfHtmlSaveOptions? options = null, Encoding? encoding = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        byte[] data = document.ToHtmlBytes(options, encoding);
        stream.Write(data, 0, data.Length);
    }

    /// <summary>Loads semantic HTML into an RTF document model.</summary>
    public static RtfDocument LoadRtfFromHtml(this string html, RtfHtmlReadOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return RtfHtmlReader.Read(html, options ?? new RtfHtmlReadOptions());
    }

    /// <summary>Loads encoded semantic HTML bytes into an RTF document model.</summary>
    public static RtfDocument LoadRtfFromHtml(this byte[] htmlBytes, RtfHtmlReadOptions? options = null, Encoding? encoding = null) {
        if (htmlBytes == null) {
            throw new ArgumentNullException(nameof(htmlBytes));
        }

        return (encoding ?? Encoding.UTF8).GetString(htmlBytes).LoadRtfFromHtml(options);
    }

    /// <summary>Reads semantic HTML from a stream into an RTF document model.</summary>
    public static RtfDocument LoadRtfFromHtml(this Stream htmlStream, RtfHtmlReadOptions? options = null, Encoding? encoding = null) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        using (var reader = new StreamReader(htmlStream, encoding ?? Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true)) {
            return reader.ReadToEnd().LoadRtfFromHtml(options);
        }
    }

    /// <summary>Converts semantic HTML to RTF text.</summary>
    public static string ToRtf(this string html, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return html.LoadRtfFromHtml(readOptions).ToRtf(writeOptions);
    }

    /// <summary>Converts semantic HTML to encoded RTF bytes.</summary>
    public static byte[] ToRtfBytes(this string html, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return html.LoadRtfFromHtml(readOptions).ToBytes(writeOptions, encoding);
    }

    /// <summary>Converts semantic HTML to an encoded RTF memory stream.</summary>
    public static MemoryStream ToRtfMemoryStream(this string html, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return html.LoadRtfFromHtml(readOptions).ToMemoryStream(writeOptions, encoding);
    }

    /// <summary>Saves semantic HTML as an RTF file at the specified path.</summary>
    public static void SaveAsRtf(this string html, string path, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        html.LoadRtfFromHtml(readOptions).Save(path, writeOptions, encoding);
    }

    /// <summary>Saves semantic HTML as RTF to a stream without closing or rewinding the stream.</summary>
    public static void SaveAsRtf(this string html, Stream stream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        html.LoadRtfFromHtml(readOptions).Save(stream, writeOptions, encoding);
    }
}
