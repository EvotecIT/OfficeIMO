namespace OfficeIMO.Html;

/// <summary>
/// Provides extension methods for converting between <see cref="RtfDocument"/> and semantic HTML.
/// </summary>
public static partial class HtmlRtfConverterExtensions {
    /// <summary>Converts an RTF document to semantic HTML.</summary>
    public static string ToHtml(this RtfDocument document, RtfToHtmlOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        RtfToHtmlOptions effectiveOptions = options ?? new RtfToHtmlOptions();
        if (effectiveOptions.PreferEncapsulatedHtml &&
            document.HtmlEncapsulation != null &&
            !string.IsNullOrWhiteSpace(document.HtmlEncapsulation.Html)) {
            var importOptions = HtmlToRtfOptions.CreateUntrustedHtmlProfile();
            importOptions.UrlPolicy = effectiveOptions.GetUrlPolicy().Clone();
            RtfDocument encapsulatedDocument = document.HtmlEncapsulation.Html.ToRtfDocument(importOptions);
            foreach (HtmlRtfConversionDiagnostic diagnostic in importOptions.Diagnostics) {
                effectiveOptions.AddDiagnostic(diagnostic.Code, diagnostic.Message, diagnostic.Source, severity: diagnostic.Severity);
            }

            effectiveOptions.AddDiagnostic(
                "RtfHtmlEncapsulatedHtmlUsed",
                "Outlook/Exchange encapsulated HTML was used instead of the RTF plain-text fallback.",
                severity: HtmlRtfConversionDiagnosticSeverity.Info);
            return RtfHtmlWriter.Write(encapsulatedDocument, effectiveOptions);
        }

        return RtfHtmlWriter.Write(document, effectiveOptions);
    }

    /// <summary>Converts an RTF document to encoded semantic HTML bytes.</summary>
    public static byte[] ToHtmlBytes(this RtfDocument document, RtfToHtmlOptions? options = null, Encoding? encoding = null) {
        return (encoding ?? Encoding.UTF8).GetBytes(document.ToHtml(options));
    }

    /// <summary>Converts an RTF document to an encoded semantic HTML memory stream.</summary>
    public static MemoryStream ToHtmlMemoryStream(this RtfDocument document, RtfToHtmlOptions? options = null, Encoding? encoding = null) {
        return new MemoryStream(document.ToHtmlBytes(options, encoding), writable: false);
    }

    /// <summary>Saves an RTF document as semantic HTML at the specified path.</summary>
    public static void SaveAsHtml(this RtfDocument document, string path, RtfToHtmlOptions? options = null, Encoding? encoding = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        File.WriteAllText(path, document.ToHtml(options), encoding ?? Encoding.UTF8);
    }

    /// <summary>Saves an RTF document as semantic HTML to a writable stream.</summary>
    public static void SaveAsHtml(this RtfDocument document, Stream stream, RtfToHtmlOptions? options = null, Encoding? encoding = null) {
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
    public static RtfDocument ToRtfDocument(this string html, HtmlToRtfOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return RtfHtmlReader.Read(html, options ?? new HtmlToRtfOptions());
    }

    /// <summary>Loads a prepared shared HTML conversion document into an RTF document model without reparsing.</summary>
    public static RtfDocument ToRtfDocument(this HtmlConversionDocument document, HtmlToRtfOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return RtfHtmlReader.Read(document.DocumentForConversion, options ?? new HtmlToRtfOptions());
    }

    /// <summary>Loads encoded semantic HTML bytes into an RTF document model.</summary>
    public static RtfDocument ToRtfDocument(this byte[] htmlBytes, HtmlToRtfOptions? options = null, Encoding? encoding = null) {
        if (htmlBytes == null) {
            throw new ArgumentNullException(nameof(htmlBytes));
        }

        return (encoding ?? Encoding.UTF8).GetString(htmlBytes).ToRtfDocument(options);
    }

    /// <summary>Reads semantic HTML from a stream into an RTF document model.</summary>
    public static RtfDocument ToRtfDocument(this Stream htmlStream, HtmlToRtfOptions? options = null, Encoding? encoding = null) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        using (var reader = new StreamReader(htmlStream, encoding ?? Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true)) {
            return reader.ReadToEnd().ToRtfDocument(options);
        }
    }

    /// <summary>Converts semantic HTML to RTF text.</summary>
    public static string ToRtf(this string html, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return html.ToRtfDocument(readOptions).ToRtf(writeOptions);
    }

    /// <summary>Converts semantic HTML to encoded RTF bytes.</summary>
    public static byte[] ToRtfBytes(this string html, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return html.ToRtfDocument(readOptions).ToBytes(writeOptions, encoding);
    }

    /// <summary>Converts semantic HTML to an encoded RTF memory stream.</summary>
    public static MemoryStream ToRtfMemoryStream(this string html, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return html.ToRtfDocument(readOptions).ToMemoryStream(writeOptions, encoding);
    }

    /// <summary>Saves semantic HTML as an RTF file at the specified path.</summary>
    public static void SaveAsRtf(this string html, string path, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        html.ToRtfDocument(readOptions).Save(path, writeOptions, encoding);
    }

    /// <summary>Saves semantic HTML as RTF to a stream without closing or rewinding the stream.</summary>
    public static void SaveAsRtf(this string html, Stream stream, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        html.ToRtfDocument(readOptions).Save(stream, writeOptions, encoding);
    }
}
