using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.AsciiDoc;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.AsciiDoc.Pdf;

/// <summary>Converts native AsciiDoc documents through the loss-aware Markdown projection to first-party PDFs.</summary>
public static class AsciiDocPdfConverterExtensions {
    /// <summary>Converts an AsciiDoc document to a first-party PDF document.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this AsciiDocDocument document, AsciiDocPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Value;

    /// <summary>Converts an AsciiDoc document and combines parser, semantic-projection, and PDF diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this AsciiDocDocument document, AsciiDocPdfSaveOptions? options = null) =>
        AsciiDocPdfConversionEngine.Convert(document, options);

    /// <summary>Converts an AsciiDoc document to PDF bytes.</summary>
    public static byte[] ToPdf(this AsciiDocDocument document, AsciiDocPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Saves an AsciiDoc document as PDF and returns combined conversion diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(this AsciiDocDocument document, string path, AsciiDocPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(path);

    /// <summary>Writes an AsciiDoc document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(this AsciiDocDocument document, Stream stream, AsciiDocPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(stream);

    /// <summary>Attempts to save an AsciiDoc PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this AsciiDocDocument document, string path, AsciiDocPdfSaveOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).TrySave(path); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Attempts to write an AsciiDoc PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this AsciiDocDocument document, Stream stream, AsciiDocPdfSaveOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).TrySave(stream); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously writes an AsciiDoc PDF to a path.</summary>
    public static Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(this AsciiDocDocument document, string path, AsciiDocPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return document.ToPdfDocumentResult(options).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes an AsciiDoc PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(this AsciiDocDocument document, Stream stream, AsciiDocPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return document.ToPdfDocumentResult(options).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Converts synchronously, then attempts to asynchronously save an AsciiDoc PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this AsciiDocDocument document, string path, AsciiDocPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await document.ToPdfDocumentResult(options).TrySaveAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Converts synchronously, then attempts to asynchronously write an AsciiDoc PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this AsciiDocDocument document, Stream stream, AsciiDocPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await document.ToPdfDocumentResult(options).TrySaveAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }
}
