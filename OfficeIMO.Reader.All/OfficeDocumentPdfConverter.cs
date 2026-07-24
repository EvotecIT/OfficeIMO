using System.IO;
using System.Threading;
using OfficeIMO.Epub;
using OfficeIMO.Pdf;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Visio;

namespace OfficeIMO.Reader.All;

/// <summary>
/// Direct Email, EPUB, and Visio PDF façades over the shared Reader normalization and PDF projection owners.
/// </summary>
public static class OfficeDocumentPdfConverter {
    /// <summary>Converts an Email artifact such as EML, MSG, OFT, or TNEF to a searchable PDF.</summary>
    public static PdfDocumentConversionResult EmailToPdf(
        Stream source,
        string sourceName = "message.eml",
        ReaderPdfProjectionOptions? pdfOptions = null,
        ReaderOptions? readerOptions = null,
        ReaderEmailOptions? emailOptions = null,
        CancellationToken cancellationToken = default) {
        return Convert(
            new OfficeDocumentReaderBuilder().AddEmailHandler(emailOptions).Build(),
            source,
            sourceName,
            pdfOptions,
            readerOptions,
            cancellationToken);
    }

    /// <summary>Converts an Email artifact such as EML, MSG, OFT, or TNEF to a searchable PDF.</summary>
    public static PdfDocumentConversionResult EmailToPdf(
        string path,
        ReaderPdfProjectionOptions? pdfOptions = null,
        ReaderOptions? readerOptions = null,
        ReaderEmailOptions? emailOptions = null,
        CancellationToken cancellationToken = default) {
        return Convert(
            new OfficeDocumentReaderBuilder().AddEmailHandler(emailOptions).Build(),
            path,
            pdfOptions,
            readerOptions,
            cancellationToken);
    }

    /// <summary>Converts an EPUB package to a searchable PDF while retaining chapter order and policy diagnostics.</summary>
    public static PdfDocumentConversionResult EpubToPdf(
        Stream source,
        string sourceName = "book.epub",
        ReaderPdfProjectionOptions? pdfOptions = null,
        ReaderOptions? readerOptions = null,
        EpubReadOptions? epubOptions = null,
        CancellationToken cancellationToken = default) {
        return Convert(
            new OfficeDocumentReaderBuilder().AddEpubHandler(epubOptions).Build(),
            source,
            sourceName,
            pdfOptions,
            readerOptions,
            cancellationToken);
    }

    /// <summary>Converts an EPUB package to a searchable PDF while retaining chapter order and policy diagnostics.</summary>
    public static PdfDocumentConversionResult EpubToPdf(
        string path,
        ReaderPdfProjectionOptions? pdfOptions = null,
        ReaderOptions? readerOptions = null,
        EpubReadOptions? epubOptions = null,
        CancellationToken cancellationToken = default) {
        return Convert(
            new OfficeDocumentReaderBuilder().AddEpubHandler(epubOptions).Build(),
            path,
            pdfOptions,
            readerOptions,
            cancellationToken);
    }

    /// <summary>Converts a Visio package to a searchable PDF with explicit preview or semantic-fallback evidence.</summary>
    public static PdfDocumentConversionResult VisioToPdf(
        Stream source,
        string sourceName = "diagram.vsdx",
        ReaderPdfProjectionOptions? pdfOptions = null,
        ReaderOptions? readerOptions = null,
        ReaderVisioOptions? visioOptions = null,
        CancellationToken cancellationToken = default) {
        return Convert(
            new OfficeDocumentReaderBuilder().AddVisioHandler(visioOptions).Build(),
            source,
            sourceName,
            pdfOptions,
            readerOptions,
            cancellationToken);
    }

    /// <summary>Converts a Visio package to a searchable PDF with explicit preview or semantic-fallback evidence.</summary>
    public static PdfDocumentConversionResult VisioToPdf(
        string path,
        ReaderPdfProjectionOptions? pdfOptions = null,
        ReaderOptions? readerOptions = null,
        ReaderVisioOptions? visioOptions = null,
        CancellationToken cancellationToken = default) {
        return Convert(
            new OfficeDocumentReaderBuilder().AddVisioHandler(visioOptions).Build(),
            path,
            pdfOptions,
            readerOptions,
            cancellationToken);
    }

    private static PdfDocumentConversionResult Convert(
        OfficeDocumentReader reader,
        Stream source,
        string sourceName,
        ReaderPdfProjectionOptions? pdfOptions,
        ReaderOptions? readerOptions,
        CancellationToken cancellationToken) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (string.IsNullOrWhiteSpace(sourceName)) throw new ArgumentException("Source name is required for format routing.", nameof(sourceName));
        OfficeDocumentReadResult normalized = reader.ReadDocument(source, sourceName, readerOptions, cancellationToken);
        return normalized.ToPdfDocumentResult(pdfOptions);
    }

    private static PdfDocumentConversionResult Convert(
        OfficeDocumentReader reader,
        string path,
        ReaderPdfProjectionOptions? pdfOptions,
        ReaderOptions? readerOptions,
        CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Source path is required.", nameof(path));
        OfficeDocumentReadResult normalized = reader.ReadDocument(path, readerOptions, cancellationToken);
        return normalized.ToPdfDocumentResult(pdfOptions);
    }
}
