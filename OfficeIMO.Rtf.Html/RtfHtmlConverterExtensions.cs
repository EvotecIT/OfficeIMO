namespace OfficeIMO.Rtf.Html;

/// <summary>
/// Provides extension methods for converting between <see cref="RtfDocument"/> and semantic HTML.
/// </summary>
public static class RtfHtmlConverterExtensions {
    /// <summary>Converts an RTF document to semantic HTML.</summary>
    public static string ToHtml(this RtfDocument document, RtfHtmlSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        return RtfHtmlWriter.Write(document, options ?? new RtfHtmlSaveOptions());
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

        byte[] data = (encoding ?? Encoding.UTF8).GetBytes(document.ToHtml(options));
        stream.Write(data, 0, data.Length);
    }

    /// <summary>Converts semantic HTML to an RTF document model.</summary>
    public static RtfDocument ToRtfDocumentFromHtml(this string html, RtfHtmlReadOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return RtfHtmlReader.Read(html, options ?? new RtfHtmlReadOptions());
    }

    /// <summary>Reads semantic HTML from a stream and converts it to an RTF document model.</summary>
    public static RtfDocument ToRtfDocumentFromHtml(this Stream htmlStream, RtfHtmlReadOptions? options = null, Encoding? encoding = null) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        using (var reader = new StreamReader(htmlStream, encoding ?? Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true)) {
            return reader.ReadToEnd().ToRtfDocumentFromHtml(options);
        }
    }

    /// <summary>Saves semantic HTML as an RTF file at the specified path.</summary>
    public static void SaveAsRtfFromHtml(this string html, string path, RtfHtmlReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        html.ToRtfDocumentFromHtml(readOptions).Save(path, writeOptions, encoding);
    }
}
