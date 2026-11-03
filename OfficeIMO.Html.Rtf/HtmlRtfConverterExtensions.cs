using OfficeIMO.Core.Internal;

namespace OfficeIMO.Html;

/// <summary>
/// Provides extension methods for converting between <see cref="RtfDocument"/> and semantic HTML.
/// </summary>
public static partial class HtmlRtfConverterExtensions {
    /// <summary>Converts an RTF document to semantic HTML.</summary>
    public static string ToHtml(this RtfDocument document, RtfHtmlExportOptions? options = null) {
        return document.ToHtmlResult(options).RequireValue();
    }

    private static string ToHtmlCore(RtfDocument document, RtfHtmlExportOptions effectiveOptions) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        HtmlConversionProfileContracts.Get(effectiveOptions.SharedProfile);
        effectiveOptions.DocumentOutput.Validate();

        if (effectiveOptions.PreferEncapsulatedHtml &&
            document.IsHtmlEncapsulationCurrent &&
            document.HtmlEncapsulation != null &&
            !string.IsNullOrWhiteSpace(document.HtmlEncapsulation.Html)) {
            var importOptions = HtmlRtfExportOptions.UntrustedDefaults;
            importOptions.UrlPolicy = effectiveOptions.GetUrlPolicy().Clone();
            HtmlConversionDocument encapsulatedHtml = HtmlConversionDocument.Parse(
                document.HtmlEncapsulation.Html,
                new HtmlConversionDocumentOptions {
                    Profile = HtmlConversionProfile.Document,
                    Trust = HtmlInputTrust.Untrusted
                });
            effectiveOptions.HtmlDiagnostics.AddRange(encapsulatedHtml.ResourceManifest.Diagnostics);
            HtmlToRtfResult imported = encapsulatedHtml.ToRtfDocumentResult(importOptions);
            foreach (HtmlRtfConversionDiagnostic diagnostic in imported.RtfDiagnostics) {
                effectiveOptions.AddDiagnostic(diagnostic.Code, diagnostic.Message, diagnostic.Source, severity: diagnostic.Severity, action: diagnostic.Action);
            }

            effectiveOptions.AddDiagnostic(
                "RtfHtmlEncapsulatedHtmlUsed",
                "Outlook/Exchange encapsulated HTML was used instead of the RTF plain-text fallback.",
                severity: HtmlRtfConversionDiagnosticSeverity.Info);
            return RtfHtmlWriter.Write(imported.RequireValue(), effectiveOptions);
        }

        if (effectiveOptions.PreferEncapsulatedHtml && document.HtmlEncapsulation != null) {
            effectiveOptions.AddDiagnostic("RtfHtmlEncapsulationInvalidated",
                "Semantic edits invalidated the original encapsulated HTML; current RTF content was converted.",
                severity: HtmlRtfConversionDiagnosticSeverity.Warning, action: RtfConversionAction.Omitted);
        }
        effectiveOptions.ConversionReport.Merge(document.GetStyleConversionDiagnostics());
        return RtfHtmlWriter.Write(document, effectiveOptions);
    }

    /// <summary>Converts an RTF document to encoded semantic HTML bytes.</summary>
    public static byte[] ToHtmlBytes(this RtfDocument document, RtfHtmlExportOptions? options = null, Encoding? encoding = null) {
        return (encoding ?? Encoding.UTF8).GetBytes(document.ToHtml(options));
    }

    /// <summary>Converts an RTF document to an encoded semantic HTML memory stream.</summary>
    public static MemoryStream ToHtmlStream(this RtfDocument document, RtfHtmlExportOptions? options = null, Encoding? encoding = null) {
        return new MemoryStream(document.ToHtmlBytes(options, encoding), writable: false);
    }

    /// <summary>Saves an RTF document as semantic HTML at the specified path.</summary>
    public static void SaveAsHtml(this RtfDocument document, string path, RtfHtmlExportOptions? options = null, Encoding? encoding = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        OfficeFileCommit.WriteAllBytes(path, (encoding ?? Encoding.UTF8).GetBytes(document.ToHtml(options)));
    }

    /// <summary>Saves an RTF document as semantic HTML to a writable stream.</summary>
    public static void SaveAsHtml(this RtfDocument document, Stream stream, RtfHtmlExportOptions? options = null, Encoding? encoding = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        OfficeStreamWriter.WriteAllBytes(stream, document.ToHtmlBytes(options, encoding));
    }

    /// <summary>Loads a prepared shared HTML conversion document into an RTF document model without reparsing.</summary>
    public static RtfDocument ToRtfDocument(this HtmlConversionDocument document, HtmlRtfExportOptions? options = null) {
        return document.ToRtfDocumentResult(options).RequireValue();
    }

    /// <summary>Converts semantic HTML to RTF text.</summary>
    public static string ToRtf(this HtmlConversionDocument document, HtmlRtfExportOptions? readOptions = null, RtfSaveOptions? writeOptions = null) {
        return document.ToRtfDocument(readOptions).ToRtf(writeOptions);
    }

    /// <summary>Converts semantic HTML to encoded RTF bytes.</summary>
    public static byte[] ToRtfBytes(this HtmlConversionDocument document, HtmlRtfExportOptions? readOptions = null, RtfSaveOptions? writeOptions = null, Encoding? encoding = null) {
        return document.ToRtfDocument(readOptions).ToBytes(writeOptions, encoding);
    }

    /// <summary>Converts semantic HTML to an encoded RTF memory stream.</summary>
    public static MemoryStream ToRtfStream(this HtmlConversionDocument document, HtmlRtfExportOptions? readOptions = null, RtfSaveOptions? writeOptions = null, Encoding? encoding = null) {
        return document.ToRtfDocument(readOptions).ToStream(writeOptions, encoding);
    }

    /// <summary>Saves semantic HTML as an RTF file at the specified path.</summary>
    public static void SaveAsRtf(this HtmlConversionDocument document, string path, HtmlRtfExportOptions? readOptions = null, RtfSaveOptions? writeOptions = null, Encoding? encoding = null) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        document.ToRtfDocument(readOptions).Save(path, writeOptions, encoding);
    }

    /// <summary>Saves semantic HTML as RTF to a stream without closing or rewinding the stream.</summary>
    public static void SaveAsRtf(this HtmlConversionDocument document, Stream stream, HtmlRtfExportOptions? readOptions = null, RtfSaveOptions? writeOptions = null, Encoding? encoding = null) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        document.ToRtfDocument(readOptions).Save(stream, writeOptions, encoding);
    }
}
