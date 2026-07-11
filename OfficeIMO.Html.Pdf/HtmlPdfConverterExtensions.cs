using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// First-party HTML to PDF conversion helpers.
/// </summary>
public static class HtmlPdfConverterExtensions {
    /// <summary>
    /// Converts HTML to a PDF document while asynchronously resolving resources for the direct rendered profile.
    /// Semantic and document profiles retain their existing synchronous conversion internals.
    /// </summary>
    public static async Task<PdfCore.PdfDocument> ToPdfDocumentAsync(this string html, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        options ??= new HtmlPdfSaveOptions();
        options.ResetExportState();
        cancellationToken.ThrowIfCancellationRequested();
        if (options.Profile != HtmlPdfProfile.Rendered) {
            return html.ToPdfDocument(options);
        }

        PdfCore.PdfDocument pdf = await HtmlPdfRenderedConverter.ConvertAsync(html, options, cancellationToken).ConfigureAwait(false);
        AddCurrentHtmlDiagnostics(options);
        cancellationToken.ThrowIfCancellationRequested();
        return pdf;
    }

    /// <summary>Converts HTML to PDF bytes with asynchronous resource resolution and cancellation.</summary>
    public static async Task<byte[]> SaveAsPdfAsync(this string html, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new HtmlPdfSaveOptions();
        PdfCore.PdfDocument pdf = await html.ToPdfDocumentAsync(options, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = pdf.ToBytes();
        cancellationToken.ThrowIfCancellationRequested();
        SyncSelectedProfileReport(options);
        return bytes;
    }

    /// <summary>Writes HTML as PDF to a stream with asynchronous resource resolution and cancellation.</summary>
    public static async Task SaveAsPdfAsync(this string html, Stream stream, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The output stream must be writable.", nameof(stream));
        byte[] bytes = await html.SaveAsPdfAsync(options, cancellationToken).ConfigureAwait(false);
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves HTML as a PDF file with asynchronous resource resolution and cancellation.</summary>
    public static async Task SaveAsPdfAsync(this string html, string path, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An output path is required.", nameof(path));
        byte[] bytes = await html.SaveAsPdfAsync(options, cancellationToken).ConfigureAwait(false);
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true);
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Converts HTML text to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDocument ToPdfDocument(this string html, HtmlPdfSaveOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        options ??= new HtmlPdfSaveOptions();
        options.ResetExportState();

        PdfCore.PdfDocument pdf = options.Profile switch {
            HtmlPdfProfile.Semantic => ConvertSemantic(html, options),
            HtmlPdfProfile.Document => ConvertDocument(html, options),
            HtmlPdfProfile.Rendered => HtmlPdfRenderedConverter.Convert(html, options),
            _ => throw new ArgumentOutOfRangeException(nameof(options.Profile), options.Profile, "Unsupported HTML PDF profile.")
        };
        AddCurrentHtmlDiagnostics(options);
        return pdf;
    }

    /// <summary>
    /// Converts a shared OfficeIMO HTML conversion document to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDocument ToPdfDocument(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new HtmlPdfSaveOptions();
        if (options.Profile == HtmlPdfProfile.Rendered) {
            options.ResetExportState();
            PdfCore.PdfDocument renderedPdf = HtmlPdfRenderedConverter.Convert(document.HtmlForConversion, options);
            AddCurrentHtmlDiagnostics(options);
            return renderedPdf;
        }

        bool useDocumentProfile = document.ProfileContract.Profile == HtmlConversionProfile.Document
            || document.ProfileContract.Profile == HtmlConversionProfile.HighFidelityPrint;
        options.Profile = useDocumentProfile
            ? HtmlPdfProfile.Document
            : HtmlPdfProfile.Semantic;
        if (useDocumentProfile) {
            options.WordHtmlOptions ??= HtmlToWordOptions.CreateTrustedDocumentProfile();
            options.WordPdfOptions ??= new PdfSaveOptions();
        }

        if (!useDocumentProfile) {
            string adapterHtml = HtmlActiveMediaFilter.Filter(document.HtmlForConversion, GetMediaContext(document));
            return adapterHtml.ToPdfDocument(options);
        }

        options.ResetExportState();
        PdfCore.PdfDocument pdf = ConvertDocument(document, options);
        AddCurrentHtmlDiagnostics(options);
        return pdf;
    }

    /// <summary>
    /// Converts HTML stream content to a first-party OfficeIMO PDF document model using UTF-8.
    /// </summary>
    public static PdfCore.PdfDocument ToPdfDocument(this Stream htmlStream, HtmlPdfSaveOptions? options = null) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        using var reader = new StreamReader(htmlStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
        return reader.ReadToEnd().ToPdfDocument(options);
    }

    /// <summary>
    /// Converts HTML text to a first-party OfficeIMO PDF document model with a snapshot of conversion diagnostics.
    /// </summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this string html, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        PdfCore.PdfDocument pdf = html.ToPdfDocument(options);
        return new PdfCore.PdfDocumentConversionResult(pdf, options.ConversionReport);
    }

    /// <summary>
    /// Converts a shared OfficeIMO HTML conversion document to a first-party OfficeIMO PDF document model with a snapshot of conversion diagnostics.
    /// </summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        PdfCore.PdfDocument pdf = document.ToPdfDocument(options);
        return new PdfCore.PdfDocumentConversionResult(pdf, options.ConversionReport);
    }

    /// <summary>
    /// Converts HTML stream content to a first-party OfficeIMO PDF document model with a snapshot of conversion diagnostics.
    /// </summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this Stream htmlStream, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        PdfCore.PdfDocument pdf = htmlStream.ToPdfDocument(options);
        return new PdfCore.PdfDocumentConversionResult(pdf, options.ConversionReport);
    }

    /// <summary>
    /// Converts HTML text to PDF bytes.
    /// </summary>
    public static byte[] SaveAsPdf(this string html, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        PdfCore.PdfDocument pdf = html.ToPdfDocument(options);
        byte[] bytes = pdf.ToBytes();
        SyncSelectedProfileReport(options);
        return bytes;
    }

    /// <summary>
    /// Converts a shared OfficeIMO HTML conversion document to PDF bytes.
    /// </summary>
    public static byte[] SaveAsPdf(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new HtmlPdfSaveOptions();
        PdfCore.PdfDocument pdf = document.ToPdfDocument(options);
        byte[] bytes = pdf.ToBytes();
        SyncSelectedProfileReport(options);
        return bytes;
    }

    /// <summary>
    /// Converts HTML stream content to PDF bytes using UTF-8.
    /// </summary>
    public static byte[] SaveAsPdf(this Stream htmlStream, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        PdfCore.PdfDocument pdf = htmlStream.ToPdfDocument(options);
        byte[] bytes = pdf.ToBytes();
        SyncSelectedProfileReport(options);
        return bytes;
    }

    /// <summary>
    /// Saves HTML text as a PDF file.
    /// </summary>
    public static void SaveAsPdf(this string html, string path, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        html.ToPdfDocument(options).Save(path);
        SyncSelectedProfileReport(options);
    }

    /// <summary>
    /// Saves a shared OfficeIMO HTML conversion document as a PDF file.
    /// </summary>
    public static void SaveAsPdf(this HtmlConversionDocument document, string path, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        document.ToPdfDocument(options).Save(path);
        SyncSelectedProfileReport(options);
    }

    /// <summary>
    /// Saves HTML stream content as a PDF file using UTF-8.
    /// </summary>
    public static void SaveAsPdf(this Stream htmlStream, string path, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        htmlStream.ToPdfDocument(options).Save(path);
        SyncSelectedProfileReport(options);
    }

    /// <summary>
    /// Attempts to save HTML text as a PDF file and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this string html, string path, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        try {
            PdfCore.PdfSaveResult result = html.ToPdfDocument(options).TrySave(path);
            SyncSelectedProfileReport(options);
            return result;
        } catch (Exception ex) {
            SyncSelectedProfileReport(options);
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>
    /// Attempts to save a shared OfficeIMO HTML conversion document as a PDF file and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this HtmlConversionDocument document, string path, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        try {
            PdfCore.PdfSaveResult result = document.ToPdfDocument(options).TrySave(path);
            SyncSelectedProfileReport(options);
            return result;
        } catch (Exception ex) {
            SyncSelectedProfileReport(options);
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>
    /// Attempts to save HTML stream content as a PDF file using UTF-8 and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this Stream htmlStream, string path, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        try {
            PdfCore.PdfSaveResult result = htmlStream.ToPdfDocument(options).TrySave(path);
            SyncSelectedProfileReport(options);
            return result;
        } catch (Exception ex) {
            SyncSelectedProfileReport(options);
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>
    /// Writes HTML text as PDF to a stream.
    /// </summary>
    public static void SaveAsPdf(this string html, Stream stream, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        html.ToPdfDocument(options).Save(stream);
        SyncSelectedProfileReport(options);
    }

    /// <summary>
    /// Writes a shared OfficeIMO HTML conversion document as PDF to a stream.
    /// </summary>
    public static void SaveAsPdf(this HtmlConversionDocument document, Stream stream, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        document.ToPdfDocument(options).Save(stream);
        SyncSelectedProfileReport(options);
    }

    /// <summary>
    /// Writes HTML stream content as PDF to a stream using UTF-8.
    /// </summary>
    public static void SaveAsPdf(this Stream htmlStream, Stream pdfStream, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        htmlStream.ToPdfDocument(options).Save(pdfStream);
        SyncSelectedProfileReport(options);
    }

    /// <summary>
    /// Attempts to write HTML text as PDF to a stream and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this string html, Stream stream, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        try {
            PdfCore.PdfSaveResult result = html.ToPdfDocument(options).TrySave(stream);
            SyncSelectedProfileReport(options);
            return result;
        } catch (Exception ex) {
            SyncSelectedProfileReport(options);
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    /// <summary>
    /// Attempts to write a shared OfficeIMO HTML conversion document as PDF to a stream and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this HtmlConversionDocument document, Stream stream, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        try {
            PdfCore.PdfSaveResult result = document.ToPdfDocument(options).TrySave(stream);
            SyncSelectedProfileReport(options);
            return result;
        } catch (Exception ex) {
            SyncSelectedProfileReport(options);
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    /// <summary>
    /// Attempts to write HTML stream content as PDF to a stream using UTF-8 and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this Stream htmlStream, Stream pdfStream, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        try {
            PdfCore.PdfSaveResult result = htmlStream.ToPdfDocument(options).TrySave(pdfStream);
            SyncSelectedProfileReport(options);
            return result;
        } catch (Exception ex) {
            SyncSelectedProfileReport(options);
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    private static PdfCore.PdfDocument ConvertSemantic(string html, HtmlPdfSaveOptions options) {
        MarkdownPdfSaveOptions markdownPdfOptions = options.MarkdownPdfOptions ?? new MarkdownPdfSaveOptions();
        options.MarkdownPdfOptions = markdownPdfOptions;
        PdfCore.PdfDocument pdf = html
            .LoadFromHtml(options.MarkdownHtmlOptions)
            .ToPdfDocument(markdownPdfOptions);
        options.ConversionReport.LinkReport(markdownPdfOptions.ConversionReport);
        return pdf;
    }

    private static PdfCore.PdfDocument ConvertDocument(string html, HtmlPdfSaveOptions options) {
        PdfSaveOptions wordPdfOptions = options.WordPdfOptions ?? new PdfSaveOptions();
        HtmlToWordOptions wordHtmlOptions = options.WordHtmlOptions ?? new HtmlToWordOptions();
        options.WordPdfOptions = wordPdfOptions;
        options.WordHtmlOptions = wordHtmlOptions;
        wordHtmlOptions.Diagnostics.Clear();
        wordHtmlOptions.ConversionReport.Clear();
        using WordDocument document = html.LoadFromHtml(wordHtmlOptions);
        PdfCore.PdfDocument pdf = document.ToPdfDocument(wordPdfOptions);
        options.ConversionReport.LinkReport(wordPdfOptions.ConversionReport);
        return pdf;
    }

    private static PdfCore.PdfDocument ConvertDocument(HtmlConversionDocument conversionDocument, HtmlPdfSaveOptions options) {
        PdfSaveOptions wordPdfOptions = options.WordPdfOptions ?? new PdfSaveOptions();
        HtmlToWordOptions wordHtmlOptions = options.WordHtmlOptions ?? HtmlToWordOptions.CreateTrustedDocumentProfile();
        options.WordPdfOptions = wordPdfOptions;
        options.WordHtmlOptions = wordHtmlOptions;
        wordHtmlOptions.Diagnostics.Clear();
        wordHtmlOptions.ConversionReport.Clear();
        using WordDocument document = WordHtmlConverterExtensions.LoadFromHtml(conversionDocument, wordHtmlOptions);
        PdfCore.PdfDocument pdf = document.ToPdfDocument(wordPdfOptions);
        options.ConversionReport.LinkReport(wordPdfOptions.ConversionReport);
        return pdf;
    }

    private static HtmlCssMediaContext GetMediaContext(HtmlConversionDocument document) {
        return document.ProfileContract.Profile == HtmlConversionProfile.HighFidelityPrint
            ? HtmlCssMediaContext.Print
            : HtmlCssMediaContext.Screen;
    }

    private static void SyncSelectedProfileReport(HtmlPdfSaveOptions options) {
        PdfCore.PdfConversionReport? source = options.Profile switch {
            HtmlPdfProfile.Semantic => options.MarkdownPdfOptions?.ConversionReport,
            HtmlPdfProfile.Document => options.WordPdfOptions?.ConversionReport,
            _ => null
        };

        List<PdfCore.PdfConversionWarning> warnings = source == null
            ? new List<PdfCore.PdfConversionWarning>()
            : new List<PdfCore.PdfConversionWarning>(source.Warnings);
        AddHtmlImportDiagnostics(options, warnings);
        AddHtmlRenderDiagnostics(options, warnings);
        options.ConversionReport.ClearLinkedReports();
        options.ConversionReport.Clear();
        options.ConversionReport.AddRange(warnings);
    }

    private static void AddCurrentHtmlDiagnostics(HtmlPdfSaveOptions options) {
        var warnings = new List<PdfCore.PdfConversionWarning>();
        AddHtmlImportDiagnostics(options, warnings);
        AddHtmlRenderDiagnostics(options, warnings);
        options.ConversionReport.AddRange(warnings);
    }

    private static void AddHtmlImportDiagnostics(HtmlPdfSaveOptions options, List<PdfCore.PdfConversionWarning> warnings) {
        if (options.Profile != HtmlPdfProfile.Document || options.WordHtmlOptions is null) {
            return;
        }

        foreach (HtmlDiagnostic diagnostic in options.WordHtmlOptions.ConversionReport.Diagnostics) {
            var details = string.IsNullOrWhiteSpace(diagnostic.Detail)
                ? null
                : new Dictionary<string, string> {
                    ["Detail"] = diagnostic.Detail!
                };
            warnings.Add(new PdfCore.PdfConversionWarning(
                diagnostic.Component,
                diagnostic.Code,
                diagnostic.Source ?? "html",
                diagnostic.Message,
                MapSeverity(diagnostic.Severity),
                details: details));
        }
    }

    private static void AddHtmlRenderDiagnostics(HtmlPdfSaveOptions options, List<PdfCore.PdfConversionWarning> warnings) {
        if (options.Profile != HtmlPdfProfile.Rendered || options.RenderDiagnostics == null) {
            return;
        }

        foreach (HtmlDiagnostic diagnostic in options.RenderDiagnostics.Diagnostics) {
            var details = string.IsNullOrWhiteSpace(diagnostic.Detail)
                ? null
                : new Dictionary<string, string> {
                    ["Detail"] = diagnostic.Detail!
                };
            warnings.Add(new PdfCore.PdfConversionWarning(
                diagnostic.Component,
                diagnostic.Code,
                diagnostic.Source ?? "html-render",
                diagnostic.Message,
                MapSeverity(diagnostic.Severity),
                details: details));
        }
    }

    private static PdfCore.PdfConversionWarningSeverity MapSeverity(HtmlDiagnosticSeverity severity) {
        return severity switch {
            HtmlDiagnosticSeverity.Info => PdfCore.PdfConversionWarningSeverity.Information,
            HtmlDiagnosticSeverity.Error => PdfCore.PdfConversionWarningSeverity.Error,
            _ => PdfCore.PdfConversionWarningSeverity.Warning
        };
    }
}
