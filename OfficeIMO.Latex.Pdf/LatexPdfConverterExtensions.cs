using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Latex;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Latex.Pdf;

/// <summary>Converts bounded-profile LaTeX documents through the loss-aware Markdown projection to first-party PDFs.</summary>
public static class LatexPdfConverterExtensions {
    /// <summary>Converts a LaTeX document to a first-party PDF document.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this LatexDocument document, LatexPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Value;

    /// <summary>Converts a LaTeX document and combines parser, semantic-projection, and PDF diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this LatexDocument document, LatexPdfSaveOptions? options = null) =>
        LatexPdfConversionEngine.Convert(document, options);

    /// <summary>Converts a LaTeX document to PDF bytes.</summary>
    public static byte[] ToPdf(this LatexDocument document, LatexPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Saves a LaTeX document as PDF and returns combined conversion diagnostics.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this LatexDocument document, string path, LatexPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(path);

    /// <summary>Writes a LaTeX document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this LatexDocument document, Stream stream, LatexPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(stream);

    /// <summary>Attempts to save a LaTeX PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this LatexDocument document, string path, LatexPdfSaveOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).TrySave(path); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Attempts to write a LaTeX PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this LatexDocument document, Stream stream, LatexPdfSaveOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).TrySave(stream); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously writes a LaTeX PDF to a path.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this LatexDocument document, string path, LatexPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return document.ToPdfDocumentResult(options).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes a LaTeX PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this LatexDocument document, Stream stream, LatexPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return document.ToPdfDocumentResult(options).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Converts synchronously, then attempts to asynchronously save a LaTeX PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this LatexDocument document, string path, LatexPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await document.ToPdfDocumentResult(options).TrySaveAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Converts synchronously, then attempts to asynchronously write a LaTeX PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this LatexDocument document, Stream stream, LatexPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await document.ToPdfDocumentResult(options).TrySaveAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }
}
