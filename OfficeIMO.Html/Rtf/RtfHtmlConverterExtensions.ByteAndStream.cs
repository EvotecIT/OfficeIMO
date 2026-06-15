namespace OfficeIMO.Html;

/// <content>
/// Provides byte and stream HTML-to-RTF extension overloads.
/// </content>
public static partial class RtfHtmlConverterExtensions {
    /// <summary>Converts encoded semantic HTML bytes to RTF text.</summary>
    public static string ToRtf(this byte[] htmlBytes, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        if (htmlBytes == null) {
            throw new ArgumentNullException(nameof(htmlBytes));
        }

        return htmlBytes.LoadRtfFromHtml(readOptions, encoding).ToRtf(writeOptions);
    }

    /// <summary>Reads semantic HTML from a stream and converts it to RTF text.</summary>
    public static string ToRtf(this Stream htmlStream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        return htmlStream.LoadRtfFromHtml(readOptions, encoding).ToRtf(writeOptions);
    }

    /// <summary>Converts encoded semantic HTML bytes to encoded RTF bytes.</summary>
    public static byte[] ToRtfBytes(this byte[] htmlBytes, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        if (htmlBytes == null) {
            throw new ArgumentNullException(nameof(htmlBytes));
        }

        return htmlBytes.LoadRtfFromHtml(readOptions, htmlEncoding).ToBytes(writeOptions, rtfEncoding);
    }

    /// <summary>Reads semantic HTML from a stream and converts it to encoded RTF bytes.</summary>
    public static byte[] ToRtfBytes(this Stream htmlStream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        return htmlStream.LoadRtfFromHtml(readOptions, htmlEncoding).ToBytes(writeOptions, rtfEncoding);
    }

    /// <summary>Converts encoded semantic HTML bytes to an encoded RTF memory stream.</summary>
    public static MemoryStream ToRtfMemoryStream(this byte[] htmlBytes, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        if (htmlBytes == null) {
            throw new ArgumentNullException(nameof(htmlBytes));
        }

        return htmlBytes.LoadRtfFromHtml(readOptions, htmlEncoding).ToMemoryStream(writeOptions, rtfEncoding);
    }

    /// <summary>Reads semantic HTML from a stream and converts it to an encoded RTF memory stream.</summary>
    public static MemoryStream ToRtfMemoryStream(this Stream htmlStream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        return htmlStream.LoadRtfFromHtml(readOptions, htmlEncoding).ToMemoryStream(writeOptions, rtfEncoding);
    }

    /// <summary>Saves encoded semantic HTML bytes as an RTF file at the specified path.</summary>
    public static void SaveAsRtf(this byte[] htmlBytes, string path, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        if (htmlBytes == null) {
            throw new ArgumentNullException(nameof(htmlBytes));
        }

        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        htmlBytes.LoadRtfFromHtml(readOptions, htmlEncoding).Save(path, writeOptions, rtfEncoding);
    }

    /// <summary>Reads semantic HTML from a stream and saves it as an RTF file at the specified path.</summary>
    public static void SaveAsRtf(this Stream htmlStream, string path, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        htmlStream.LoadRtfFromHtml(readOptions, htmlEncoding).Save(path, writeOptions, rtfEncoding);
    }

    /// <summary>Saves encoded semantic HTML bytes as RTF to a writable stream without closing or rewinding the stream.</summary>
    public static void SaveAsRtf(this byte[] htmlBytes, Stream stream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        if (htmlBytes == null) {
            throw new ArgumentNullException(nameof(htmlBytes));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        htmlBytes.LoadRtfFromHtml(readOptions, htmlEncoding).Save(stream, writeOptions, rtfEncoding);
    }

    /// <summary>Reads semantic HTML from a stream and saves it as RTF to a writable stream without closing or rewinding the stream.</summary>
    public static void SaveAsRtf(this Stream htmlStream, Stream stream, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? htmlEncoding = null, Encoding? rtfEncoding = null) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        htmlStream.LoadRtfFromHtml(readOptions, htmlEncoding).Save(stream, writeOptions, rtfEncoding);
    }
}
