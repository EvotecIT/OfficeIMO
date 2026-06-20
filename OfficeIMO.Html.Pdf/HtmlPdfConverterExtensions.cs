using System.Collections.Generic;
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
            _ => throw new ArgumentOutOfRangeException(nameof(options.Profile), options.Profile, "Unsupported HTML PDF profile.")
        };
        AddCurrentHtmlImportDiagnostics(options);
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
        bool useDocumentProfile = document.ProfileContract.Profile == HtmlConversionProfile.Document
            || document.ProfileContract.Profile == HtmlConversionProfile.HighFidelityPrint;
        options.Profile = useDocumentProfile
            ? HtmlPdfProfile.Document
            : HtmlPdfProfile.Semantic;
        if (useDocumentProfile) {
            options.WordHtmlOptions ??= HtmlToWordOptions.CreateTrustedDocumentProfile();
            options.WordPdfOptions ??= new PdfSaveOptions();
        }

        return document.HtmlForConversion.ToPdfDocument(options);
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
    /// Saves HTML text as a PDF file.
    /// </summary>
    public static void SaveAsPdf(this string html, string path, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        html.ToPdfDocument(options).Save(path);
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
    /// Writes HTML text as PDF to a stream.
    /// </summary>
    public static void SaveAsPdf(this string html, Stream stream, HtmlPdfSaveOptions? options = null) {
        options ??= new HtmlPdfSaveOptions();
        html.ToPdfDocument(options).Save(stream);
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
        options.ConversionReport.ClearLinkedReports();
        options.ConversionReport.Clear();
        options.ConversionReport.AddRange(warnings);
    }

    private static void AddCurrentHtmlImportDiagnostics(HtmlPdfSaveOptions options) {
        if (options.Profile != HtmlPdfProfile.Document) {
            return;
        }

        var warnings = new List<PdfCore.PdfConversionWarning>();
        AddHtmlImportDiagnostics(options, warnings);
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

    private static PdfCore.PdfConversionWarningSeverity MapSeverity(HtmlDiagnosticSeverity severity) {
        return severity switch {
            HtmlDiagnosticSeverity.Info => PdfCore.PdfConversionWarningSeverity.Information,
            HtmlDiagnosticSeverity.Error => PdfCore.PdfConversionWarningSeverity.Error,
            _ => PdfCore.PdfConversionWarningSeverity.Warning
        };
    }
}
