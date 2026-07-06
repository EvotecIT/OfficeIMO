using PdfCore = OfficeIMO.Pdf;
using System.Text;

namespace OfficeIMO.Rtf.Pdf;

/// <content>
/// Provides explicit file-source RTF to PDF conversion APIs.
/// </content>
public static partial class RtfPdfConverterExtensions {
    /// <summary>Converts an RTF file to a first-party OfficeIMO PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocumentFromRtfFile(this string path, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        return RtfDocument.Load(path, readOptions, encoding).Document.ToPdfDocument(options);
    }

    /// <summary>Converts an RTF file to PDF and returns the generated document with a snapshot of conversion diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResultFromRtfFile(this string path, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        return RtfDocument.Load(path, readOptions, encoding).Document.ToPdfDocumentResult(options);
    }

    /// <summary>Converts an RTF file to a first-party OfficeIMO PDF document model asynchronously.</summary>
    public static async Task<PdfCore.PdfDocument> ToPdfDocumentFromRtfFileAsync(this string path, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        cancellationToken.ThrowIfCancellationRequested();
        RtfReadResult readResult = await RtfDocument.LoadAsync(path, readOptions, encoding, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return readResult.Document.ToPdfDocument(options);
    }

    /// <summary>Converts an RTF file to PDF asynchronously and returns the generated document with a snapshot of conversion diagnostics.</summary>
    public static async Task<PdfCore.PdfDocumentConversionResult> ToPdfDocumentResultFromRtfFileAsync(this string path, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        cancellationToken.ThrowIfCancellationRequested();
        RtfReadResult readResult = await RtfDocument.LoadAsync(path, readOptions, encoding, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return readResult.Document.ToPdfDocumentResult(options);
    }

    /// <summary>Converts an RTF file to PDF bytes.</summary>
    public static byte[] SaveRtfFileAsPdf(string rtfPath, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null) {
        return rtfPath.ToPdfDocumentFromRtfFile(readOptions, options, encoding).ToBytes();
    }

    /// <summary>Saves an RTF file as a PDF file.</summary>
    public static void SaveRtfFileAsPdf(string rtfPath, string pdfPath, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null) {
        rtfPath.ToPdfDocumentFromRtfFile(readOptions, options, encoding).Save(pdfPath);
    }

    /// <summary>Saves an RTF file as PDF to a writable stream.</summary>
    public static void SaveRtfFileAsPdf(string rtfPath, Stream stream, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null) {
        rtfPath.ToPdfDocumentFromRtfFile(readOptions, options, encoding).Save(stream);
    }

    /// <summary>Attempts to save an RTF file as a PDF file and returns diagnostics instead of throwing.</summary>
    public static PdfCore.PdfSaveResult TrySaveRtfFileAsPdf(string rtfPath, string pdfPath, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null) {
        try {
            return rtfPath.ToPdfDocumentFromRtfFile(readOptions, options, encoding).TrySave(pdfPath);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(pdfPath, ex);
        }
    }

    /// <summary>Attempts to save an RTF file as PDF to a writable stream and returns diagnostics instead of throwing.</summary>
    public static PdfCore.PdfSaveResult TrySaveRtfFileAsPdf(string rtfPath, Stream stream, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null) {
        try {
            return rtfPath.ToPdfDocumentFromRtfFile(readOptions, options, encoding).TrySave(stream);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    /// <summary>Converts an RTF file to PDF bytes asynchronously.</summary>
    public static async Task<byte[]> SaveRtfFileAsPdfAsync(string rtfPath, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        PdfCore.PdfDocument document = await rtfPath.ToPdfDocumentFromRtfFileAsync(readOptions, options, encoding, cancellationToken).ConfigureAwait(false);
        return await SavePdfDocumentAsBytesAsync(document, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves an RTF file as a PDF file asynchronously.</summary>
    public static async Task SaveRtfFileAsPdfAsync(string rtfPath, string pdfPath, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        PdfCore.PdfDocument document = await rtfPath.ToPdfDocumentFromRtfFileAsync(readOptions, options, encoding, cancellationToken).ConfigureAwait(false);
        await document.SaveAsync(pdfPath, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves an RTF file as PDF to a writable stream asynchronously.</summary>
    public static async Task SaveRtfFileAsPdfAsync(string rtfPath, Stream stream, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        PdfCore.PdfDocument document = await rtfPath.ToPdfDocumentFromRtfFileAsync(readOptions, options, encoding, cancellationToken).ConfigureAwait(false);
        await document.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Attempts to save an RTF file as a PDF file asynchronously and returns diagnostics instead of throwing.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveRtfFileAsPdfAsync(string rtfPath, string pdfPath, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        try {
            PdfCore.PdfDocument document = await rtfPath.ToPdfDocumentFromRtfFileAsync(readOptions, options, encoding, cancellationToken).ConfigureAwait(false);
            return await document.TrySaveAsync(pdfPath, cancellationToken).ConfigureAwait(false);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(pdfPath, ex);
        }
    }

    /// <summary>Attempts to save an RTF file as PDF to a writable stream asynchronously and returns diagnostics instead of throwing.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveRtfFileAsPdfAsync(string rtfPath, Stream stream, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        try {
            PdfCore.PdfDocument document = await rtfPath.ToPdfDocumentFromRtfFileAsync(readOptions, options, encoding, cancellationToken).ConfigureAwait(false);
            return await document.TrySaveAsync(stream, cancellationToken).ConfigureAwait(false);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }
}
