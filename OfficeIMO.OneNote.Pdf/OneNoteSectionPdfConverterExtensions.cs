using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

/// <summary>Converts offline OneNote sections to semantic PDF documents.</summary>
public static class OneNoteSectionPdfConverterExtensions {
    /// <summary>Converts a section to a first-party PDF document.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this OneNoteSection section, OneNotePdfSaveOptions? options = null) =>
        section.ToPdfDocumentResult(options).Value;

    /// <summary>Converts a section and returns explicit source, projection, and PDF diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this OneNoteSection section, OneNotePdfSaveOptions? options = null) =>
        OneNotePdfConversionEngine.Convert(section, options);

    /// <summary>Converts a section to PDF bytes.</summary>
    public static byte[] ToPdf(this OneNoteSection section, OneNotePdfSaveOptions? options = null) =>
        section.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Saves a section as PDF and returns conversion diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(this OneNoteSection section, string path, OneNotePdfSaveOptions? options = null) =>
        section.ToPdfDocumentResult(options).Save(path);

    /// <summary>Writes a section as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(this OneNoteSection section, Stream stream, OneNotePdfSaveOptions? options = null) =>
        section.ToPdfDocumentResult(options).Save(stream);

    /// <summary>Attempts to save a section as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this OneNoteSection section, string path, OneNotePdfSaveOptions? options = null) {
        try { return section.ToPdfDocumentResult(options).TrySave(path); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Attempts to write a section as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this OneNoteSection section, Stream stream, OneNotePdfSaveOptions? options = null) {
        try { return section.ToPdfDocumentResult(options).TrySave(stream); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously writes a section PDF to a path.</summary>
    public static Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(this OneNoteSection section, string path, OneNotePdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return section.ToPdfDocumentResult(options).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes a section PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(this OneNoteSection section, Stream stream, OneNotePdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return section.ToPdfDocumentResult(options).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Converts synchronously, then attempts to asynchronously save a section PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this OneNoteSection section, string path, OneNotePdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await section.ToPdfDocumentResult(options).TrySaveAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Converts synchronously, then attempts to asynchronously write a section PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this OneNoteSection section, Stream stream, OneNotePdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await section.ToPdfDocumentResult(options).TrySaveAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }
}
