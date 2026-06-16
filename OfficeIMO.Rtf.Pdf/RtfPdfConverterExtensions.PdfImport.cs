using PdfCore = OfficeIMO.Pdf;
using System.Text;

namespace OfficeIMO.Rtf.Pdf;

/// <content>
/// Provides explicit PDF-source conversion APIs for semantic PDF to RTF import.
/// </content>
public static partial class RtfPdfConverterExtensions {
    /// <summary>Converts a first-party PDF read model into an RTF document using semantic text extraction.</summary>
    public static RtfDocument ToRtfDocument(this PdfCore.PdfReadDocument document, PdfRtfReadOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        return PdfRtfConverter.Convert(document, options);
    }

    /// <summary>Converts a first-party logical PDF read model into an RTF document using semantic text extraction.</summary>
    public static RtfDocument ToRtfDocument(this PdfCore.PdfLogicalDocument document, PdfRtfReadOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        return PdfRtfConverter.Convert(document, options);
    }

    /// <summary>Reads source PDF bytes and converts parser-supported semantic content to an RTF document.</summary>
    public static RtfDocument ToRtfDocumentFromPdf(this byte[] pdfBytes, PdfRtfReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
        if (pdfBytes == null) {
            throw new ArgumentNullException(nameof(pdfBytes));
        }

        return PdfCore.PdfReadDocument.Load(pdfBytes, readOptions).ToRtfDocument(options);
    }

    /// <summary>Reads a PDF stream from the current position and converts parser-supported semantic content to an RTF document.</summary>
    public static RtfDocument ToRtfDocumentFromPdf(this Stream pdfStream, PdfRtfReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
        if (pdfStream == null) {
            throw new ArgumentNullException(nameof(pdfStream));
        }

        return PdfCore.PdfReadDocument.Load(pdfStream, readOptions).ToRtfDocument(options);
    }

    /// <summary>Reads a PDF file and converts parser-supported semantic content to an RTF document.</summary>
    public static RtfDocument ToRtfDocumentFromPdfFile(this string pdfPath, PdfRtfReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
        if (pdfPath == null) {
            throw new ArgumentNullException(nameof(pdfPath));
        }

        return PdfCore.PdfReadDocument.Load(pdfPath, readOptions).ToRtfDocument(options);
    }

    /// <summary>Reads source PDF bytes and returns serialized RTF text.</summary>
    public static string ToRtfFromPdf(this byte[] pdfBytes, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        return pdfBytes.ToRtfDocumentFromPdf(readOptions, pdfReadOptions).ToRtf(writeOptions);
    }

    /// <summary>Reads a PDF stream and returns serialized RTF text.</summary>
    public static string ToRtfFromPdf(this Stream pdfStream, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        return pdfStream.ToRtfDocumentFromPdf(readOptions, pdfReadOptions).ToRtf(writeOptions);
    }

    /// <summary>Reads a PDF file and returns serialized RTF text.</summary>
    public static string ToRtfFromPdfFile(this string pdfPath, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        return pdfPath.ToRtfDocumentFromPdfFile(readOptions, pdfReadOptions).ToRtf(writeOptions);
    }

    /// <summary>Reads source PDF bytes and returns encoded RTF bytes.</summary>
    public static byte[] ToRtfBytesFromPdf(this byte[] pdfBytes, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        return pdfBytes.ToRtfDocumentFromPdf(readOptions, pdfReadOptions).ToBytes(writeOptions, encoding);
    }

    /// <summary>Reads a PDF stream and returns encoded RTF bytes.</summary>
    public static byte[] ToRtfBytesFromPdf(this Stream pdfStream, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        return pdfStream.ToRtfDocumentFromPdf(readOptions, pdfReadOptions).ToBytes(writeOptions, encoding);
    }

    /// <summary>Reads a PDF file and returns encoded RTF bytes.</summary>
    public static byte[] ToRtfBytesFromPdfFile(this string pdfPath, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        return pdfPath.ToRtfDocumentFromPdfFile(readOptions, pdfReadOptions).ToBytes(writeOptions, encoding);
    }

    /// <summary>Reads source PDF bytes and saves semantic RTF output to a file.</summary>
    public static void SavePdfAsRtf(this byte[] pdfBytes, string rtfPath, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        pdfBytes.ToRtfDocumentFromPdf(readOptions, pdfReadOptions).Save(rtfPath, writeOptions, encoding);
    }

    /// <summary>Reads a PDF stream and saves semantic RTF output to a file.</summary>
    public static void SavePdfAsRtf(this Stream pdfStream, string rtfPath, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        pdfStream.ToRtfDocumentFromPdf(readOptions, pdfReadOptions).Save(rtfPath, writeOptions, encoding);
    }

    /// <summary>Reads a PDF file and saves semantic RTF output to a file.</summary>
    public static void SavePdfFileAsRtf(string pdfPath, string rtfPath, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        pdfPath.ToRtfDocumentFromPdfFile(readOptions, pdfReadOptions).Save(rtfPath, writeOptions, encoding);
    }

    /// <summary>Reads source PDF bytes and saves semantic RTF output to a writable stream.</summary>
    public static void SavePdfAsRtf(this byte[] pdfBytes, Stream rtfStream, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        if (rtfStream == null) {
            throw new ArgumentNullException(nameof(rtfStream));
        }

        pdfBytes.ToRtfDocumentFromPdf(readOptions, pdfReadOptions).Save(rtfStream, writeOptions, encoding);
    }

    /// <summary>Reads a PDF stream and saves semantic RTF output to a writable stream.</summary>
    public static void SavePdfAsRtf(this Stream pdfStream, Stream rtfStream, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        if (rtfStream == null) {
            throw new ArgumentNullException(nameof(rtfStream));
        }

        pdfStream.ToRtfDocumentFromPdf(readOptions, pdfReadOptions).Save(rtfStream, writeOptions, encoding);
    }

    /// <summary>Reads a PDF file and saves semantic RTF output to a writable stream.</summary>
    public static void SavePdfFileAsRtf(string pdfPath, Stream rtfStream, PdfRtfReadOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, PdfCore.PdfReadOptions? pdfReadOptions = null) {
        if (rtfStream == null) {
            throw new ArgumentNullException(nameof(rtfStream));
        }

        pdfPath.ToRtfDocumentFromPdfFile(readOptions, pdfReadOptions).Save(rtfStream, writeOptions, encoding);
    }
}
