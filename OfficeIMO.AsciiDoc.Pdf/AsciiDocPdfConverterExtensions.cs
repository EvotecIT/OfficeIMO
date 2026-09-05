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
    public static PdfCore.PdfDocument ToPdfDocument(this AsciiDocDocument document, AsciiDocToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Value;

    /// <summary>Converts an AsciiDoc document and combines parser, semantic-projection, and PDF diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this AsciiDocDocument document, AsciiDocToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        AsciiDocPdfConversionEngine.Convert(document, options, cancellationToken);

    /// <summary>Converts an AsciiDoc document to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this AsciiDocDocument document, AsciiDocToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).ToBytes(cancellationToken);

    /// <summary>Saves an AsciiDoc document as PDF and returns combined conversion diagnostics.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this AsciiDocDocument document, string path, AsciiDocToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);

    /// <summary>Writes an AsciiDoc document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this AsciiDocDocument document, Stream stream, AsciiDocToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(stream, cancellationToken);

    /// <summary>Attempts to save an AsciiDoc PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this AsciiDocDocument document, string path, AsciiDocToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return document.ToPdfDocumentResult(options, cancellationToken).SaveResult(path, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Attempts to write an AsciiDoc PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this AsciiDocDocument document, Stream stream, AsciiDocToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return document.ToPdfDocumentResult(options, cancellationToken).SaveResult(stream, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously writes an AsciiDoc PDF to a path.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this AsciiDocDocument document, string path, AsciiDocToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return document.ToPdfDocumentResult(options, cancellationToken).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes an AsciiDoc PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this AsciiDocDocument document, Stream stream, AsciiDocToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return document.ToPdfDocumentResult(options, cancellationToken).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Converts synchronously, then attempts to asynchronously save an AsciiDoc PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this AsciiDocDocument document, string path, AsciiDocToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await document.ToPdfDocumentResult(options, cancellationToken).SaveResultAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Converts synchronously, then attempts to asynchronously write an AsciiDoc PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this AsciiDocDocument document, Stream stream, AsciiDocToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await document.ToPdfDocumentResult(options, cancellationToken).SaveResultAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }
}
