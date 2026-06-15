namespace OfficeIMO.Rtf.Html;

/// <summary>
/// Converts between semantic HTML and the dependency-free OfficeIMO RTF document model.
/// </summary>
public static class RtfHtmlConverter {
    /// <summary>Converts semantic HTML to an RTF document model.</summary>
    public static RtfDocument FromHtml(string html, RtfHtmlReadOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return RtfHtmlReader.Read(html, options ?? new RtfHtmlReadOptions());
    }

    /// <summary>Converts encoded semantic HTML bytes to an RTF document model.</summary>
    public static RtfDocument FromHtml(byte[] htmlBytes, RtfHtmlReadOptions? options = null, Encoding? encoding = null) {
        if (htmlBytes == null) {
            throw new ArgumentNullException(nameof(htmlBytes));
        }

        return FromHtml((encoding ?? Encoding.UTF8).GetString(htmlBytes), options);
    }

    /// <summary>Reads semantic HTML from a stream and converts it to an RTF document model.</summary>
    public static RtfDocument FromHtml(Stream htmlStream, RtfHtmlReadOptions? options = null, Encoding? encoding = null) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        using (var reader = new StreamReader(htmlStream, encoding ?? Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true)) {
            return FromHtml(reader.ReadToEnd(), options);
        }
    }

    /// <summary>Converts an RTF document model to semantic HTML.</summary>
    public static string ToHtml(RtfDocument document, RtfHtmlSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        return RtfHtmlWriter.Write(document, options ?? new RtfHtmlSaveOptions());
    }

    /// <summary>Converts an RTF document model to encoded semantic HTML bytes.</summary>
    public static byte[] ToHtmlBytes(RtfDocument document, RtfHtmlSaveOptions? options = null, Encoding? encoding = null) {
        return (encoding ?? Encoding.UTF8).GetBytes(ToHtml(document, options));
    }

    /// <summary>Converts an RTF document model to an encoded semantic HTML memory stream.</summary>
    public static MemoryStream ToHtmlMemoryStream(RtfDocument document, RtfHtmlSaveOptions? options = null, Encoding? encoding = null) {
        return new MemoryStream(ToHtmlBytes(document, options, encoding), writable: false);
    }

    /// <summary>Saves an RTF document model as semantic HTML at the specified path.</summary>
    public static void SaveAsHtml(RtfDocument document, string path, RtfHtmlSaveOptions? options = null, Encoding? encoding = null) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        File.WriteAllText(path, ToHtml(document, options), encoding ?? Encoding.UTF8);
    }

    /// <summary>Saves an RTF document model as semantic HTML to a writable stream.</summary>
    public static void SaveAsHtml(RtfDocument document, Stream stream, RtfHtmlSaveOptions? options = null, Encoding? encoding = null) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        byte[] data = ToHtmlBytes(document, options, encoding);
        stream.Write(data, 0, data.Length);
    }

    /// <summary>Converts semantic HTML directly to RTF text.</summary>
    public static string ToRtf(string html, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null) {
        return FromHtml(html, readOptions).ToRtf(writeOptions);
    }

    /// <summary>Converts encoded semantic HTML bytes directly to RTF text.</summary>
    public static string ToRtf(byte[] htmlBytes, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        return FromHtml(htmlBytes, readOptions, encoding).ToRtf(writeOptions);
    }

    /// <summary>Reads semantic HTML from a stream and converts it directly to RTF text.</summary>
    public static string ToRtf(Stream htmlStream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        return FromHtml(htmlStream, readOptions, encoding).ToRtf(writeOptions);
    }

    /// <summary>Converts semantic HTML directly to encoded RTF bytes.</summary>
    public static byte[] ToRtfBytes(string html, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        return (encoding ?? Encoding.UTF8).GetBytes(ToRtf(html, readOptions, writeOptions));
    }

    /// <summary>Converts encoded semantic HTML bytes directly to encoded RTF bytes.</summary>
    public static byte[] ToRtfBytes(byte[] htmlBytes, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        return (rtfEncoding ?? Encoding.UTF8).GetBytes(ToRtf(htmlBytes, readOptions, writeOptions, htmlEncoding));
    }

    /// <summary>Reads semantic HTML from a stream and converts it directly to encoded RTF bytes.</summary>
    public static byte[] ToRtfBytes(Stream htmlStream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        return (rtfEncoding ?? Encoding.UTF8).GetBytes(ToRtf(htmlStream, readOptions, writeOptions, htmlEncoding));
    }

    /// <summary>Converts semantic HTML directly to an encoded RTF memory stream.</summary>
    public static MemoryStream ToRtfMemoryStream(string html, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        return new MemoryStream(ToRtfBytes(html, readOptions, writeOptions, encoding), writable: false);
    }

    /// <summary>Converts encoded semantic HTML bytes directly to an encoded RTF memory stream.</summary>
    public static MemoryStream ToRtfMemoryStream(byte[] htmlBytes, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        return new MemoryStream(ToRtfBytes(htmlBytes, readOptions, writeOptions, htmlEncoding, rtfEncoding), writable: false);
    }

    /// <summary>Reads semantic HTML from a stream and converts it directly to an encoded RTF memory stream.</summary>
    public static MemoryStream ToRtfMemoryStream(Stream htmlStream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        return new MemoryStream(ToRtfBytes(htmlStream, readOptions, writeOptions, htmlEncoding, rtfEncoding), writable: false);
    }

    /// <summary>Saves semantic HTML as RTF at the specified path.</summary>
    public static void SaveAsRtf(string html, string path, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        FromHtml(html, readOptions).Save(path, writeOptions, encoding);
    }

    /// <summary>Saves encoded semantic HTML bytes as RTF at the specified path.</summary>
    public static void SaveAsRtf(byte[] htmlBytes, string path, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        FromHtml(htmlBytes, readOptions, htmlEncoding).Save(path, writeOptions, rtfEncoding);
    }

    /// <summary>Reads semantic HTML from a stream and saves it as RTF at the specified path.</summary>
    public static void SaveAsRtf(Stream htmlStream, string path, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        FromHtml(htmlStream, readOptions, htmlEncoding).Save(path, writeOptions, rtfEncoding);
    }

    /// <summary>Saves semantic HTML as RTF to a writable stream without closing or rewinding the stream.</summary>
    public static void SaveAsRtf(string html, Stream stream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        FromHtml(html, readOptions).Save(stream, writeOptions, encoding);
    }

    /// <summary>Saves encoded semantic HTML bytes as RTF to a writable stream without closing or rewinding the stream.</summary>
    public static void SaveAsRtf(byte[] htmlBytes, Stream stream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        FromHtml(htmlBytes, readOptions, htmlEncoding).Save(stream, writeOptions, rtfEncoding);
    }

    /// <summary>Reads semantic HTML from a stream and saves it as RTF to a writable stream without closing or rewinding the stream.</summary>
    public static void SaveAsRtf(Stream htmlStream, Stream stream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        FromHtml(htmlStream, readOptions, htmlEncoding).Save(stream, writeOptions, rtfEncoding);
    }
}
