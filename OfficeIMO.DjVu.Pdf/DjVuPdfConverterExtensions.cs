using OfficeIMO.Pdf;

namespace OfficeIMO.DjVu.Pdf;

/// <summary>Scanned-page PDFs with physical source geometry and a searchable stored or explicit OCR text layer.</summary>
public static class DjVuPdfConverterExtensions {
    /// <summary>Converts to a PDF model with DjVu-stage fidelity and page provenance reports.</summary>
    public static PdfDocumentConversionResult ToPdfDocumentResult(this DjVuDocument document, DjVuToPdfOptions? options = null,
        CancellationToken cancellationToken = default) => DjVuPdfConversionEngine.Convert(document, options, cancellationToken);
    /// <summary>Converts to a PDF model, optionally recognizing only absent or empty stored text with the supplied engine.</summary>
    public static Task<PdfDocumentConversionResult> ToPdfDocumentResultAsync(this DjVuDocument document, DjVuToPdfOptions? options = null,
        CancellationToken cancellationToken = default) => DjVuPdfConversionEngine.ConvertAsync(document, options, cancellationToken);
    /// <summary>Converts using stored text and returns serialized PDF bytes.</summary>
    public static byte[] ToPdfBytes(this DjVuDocument document, DjVuToPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).ToBytes(cancellationToken);
    /// <summary>Converts with optional explicit OCR and returns serialized PDF bytes.</summary>
    public static async Task<byte[]> ToPdfBytesAsync(this DjVuDocument document, DjVuToPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).ToBytes(cancellationToken);
    /// <summary>Saves a scanned-page PDF and returns combined conversion diagnostics.</summary>
    public static PdfSaveResult SaveAsPdf(this DjVuDocument document, string path, DjVuToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return document.ToPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);
    }
    /// <summary>Writes a scanned-page PDF to a caller-owned stream. Seekable output is replaced and rewound; the stream remains open.</summary>
    /// <remarks>A serialization failure can leave partial output. Hosts publishing files should stage and validate the artifact first.</remarks>
    public static PdfSaveResult SaveAsPdf(this DjVuDocument document, Stream output, DjVuToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        ValidateOutput(output);
        return document.ToPdfDocumentResult(options, cancellationToken).Save(output, cancellationToken);
    }
    /// <summary>Converts with optional explicit OCR, then asynchronously saves the PDF.</summary>
    public static async Task<PdfSaveResult> SaveAsPdfAsync(this DjVuDocument document, string path, DjVuToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).SaveAsync(path, cancellationToken).ConfigureAwait(false);
    }
    /// <summary>Converts with optional explicit OCR, then asynchronously writes to a caller-owned stream.</summary>
    public static async Task<PdfSaveResult> SaveAsPdfAsync(this DjVuDocument document, Stream output, DjVuToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        ValidateOutput(output);
        return await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).SaveAsync(output, cancellationToken).ConfigureAwait(false);
    }
    private static void ValidateOutput(Stream output) {
        if (output == null) throw new ArgumentNullException(nameof(output));
        if (!output.CanWrite) throw new ArgumentException("PDF output stream must be writable.", nameof(output));
    }
}
