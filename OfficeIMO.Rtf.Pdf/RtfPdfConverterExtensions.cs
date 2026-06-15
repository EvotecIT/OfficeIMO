using PdfCore = OfficeIMO.Pdf;
using System.Text;

namespace OfficeIMO.Rtf.Pdf;

/// <summary>
/// Provides extension methods for converting <see cref="RtfDocument"/> instances and RTF strings to PDF files.
/// </summary>
public static class RtfPdfConverterExtensions {
    /// <summary>Converts an RTF document to a first-party OfficeIMO PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this RtfDocument document, RtfPdfSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        return RtfPdfConverter.Convert(document, options);
    }

    /// <summary>Reads an RTF string and converts it to a first-party OfficeIMO PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this string rtf, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null) {
        if (rtf == null) {
            throw new ArgumentNullException(nameof(rtf));
        }

        return RtfDocument.Read(rtf, readOptions).Document.ToPdfDocument(options);
    }

    /// <summary>Reads source RTF bytes and converts them to a first-party OfficeIMO PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this byte[] rtfBytes, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null) {
        if (rtfBytes == null) {
            throw new ArgumentNullException(nameof(rtfBytes));
        }

        return RtfDocument.Load(rtfBytes, readOptions).Document.ToPdfDocument(options);
    }

    /// <summary>Reads an RTF stream from the current position and converts it to a first-party OfficeIMO PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this Stream rtfStream, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null) {
        if (rtfStream == null) {
            throw new ArgumentNullException(nameof(rtfStream));
        }

        return RtfDocument.Load(rtfStream, readOptions, encoding).Document.ToPdfDocument(options);
    }

    /// <summary>Saves an RTF document as PDF at the specified path.</summary>
    public static void SaveAsPdf(this RtfDocument document, string path, RtfPdfSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        document.ToPdfDocument(options).Save(path);
    }

    /// <summary>Saves an RTF string as PDF at the specified path.</summary>
    public static void SaveAsPdf(this string rtf, string path, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null) {
        if (rtf == null) {
            throw new ArgumentNullException(nameof(rtf));
        }

        rtf.ToPdfDocument(readOptions, options).Save(path);
    }

    /// <summary>Saves source RTF bytes as PDF at the specified path.</summary>
    public static void SaveAsPdf(this byte[] rtfBytes, string path, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null) {
        if (rtfBytes == null) {
            throw new ArgumentNullException(nameof(rtfBytes));
        }

        rtfBytes.ToPdfDocument(readOptions, options).Save(path);
    }

    /// <summary>Saves an RTF stream as PDF at the specified path.</summary>
    public static void SaveAsPdf(this Stream rtfStream, string path, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null) {
        if (rtfStream == null) {
            throw new ArgumentNullException(nameof(rtfStream));
        }

        rtfStream.ToPdfDocument(readOptions, options, encoding).Save(path);
    }

    /// <summary>Saves an RTF document as PDF to a writable stream.</summary>
    public static void SaveAsPdf(this RtfDocument document, Stream stream, RtfPdfSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        document.ToPdfDocument(options).Save(stream);
    }

    /// <summary>Saves an RTF string as PDF to a writable stream.</summary>
    public static void SaveAsPdf(this string rtf, Stream stream, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null) {
        if (rtf == null) {
            throw new ArgumentNullException(nameof(rtf));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        rtf.ToPdfDocument(readOptions, options).Save(stream);
    }

    /// <summary>Saves source RTF bytes as PDF to a writable stream.</summary>
    public static void SaveAsPdf(this byte[] rtfBytes, Stream stream, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null) {
        if (rtfBytes == null) {
            throw new ArgumentNullException(nameof(rtfBytes));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        rtfBytes.ToPdfDocument(readOptions, options).Save(stream);
    }

    /// <summary>Saves an RTF stream as PDF to a writable stream.</summary>
    public static void SaveAsPdf(this Stream rtfStream, Stream pdfStream, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null) {
        if (rtfStream == null) {
            throw new ArgumentNullException(nameof(rtfStream));
        }

        if (pdfStream == null) {
            throw new ArgumentNullException(nameof(pdfStream));
        }

        rtfStream.ToPdfDocument(readOptions, options, encoding).Save(pdfStream);
    }

    /// <summary>Saves an RTF document as PDF and returns the generated bytes.</summary>
    public static byte[] SaveAsPdf(this RtfDocument document, RtfPdfSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        return document.ToPdfDocument(options).ToBytes();
    }

    /// <summary>Saves an RTF string as PDF and returns the generated bytes.</summary>
    public static byte[] SaveAsPdf(this string rtf, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null) {
        if (rtf == null) {
            throw new ArgumentNullException(nameof(rtf));
        }

        return rtf.ToPdfDocument(readOptions, options).ToBytes();
    }

    /// <summary>Saves source RTF bytes as PDF and returns the generated bytes.</summary>
    public static byte[] SaveAsPdf(this byte[] rtfBytes, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null) {
        if (rtfBytes == null) {
            throw new ArgumentNullException(nameof(rtfBytes));
        }

        return rtfBytes.ToPdfDocument(readOptions, options).ToBytes();
    }

    /// <summary>Saves an RTF stream as PDF and returns the generated bytes.</summary>
    public static byte[] SaveAsPdf(this Stream rtfStream, RtfReadOptions? readOptions = null, RtfPdfSaveOptions? options = null, Encoding? encoding = null) {
        if (rtfStream == null) {
            throw new ArgumentNullException(nameof(rtfStream));
        }

        return rtfStream.ToPdfDocument(readOptions, options, encoding).ToBytes();
    }

    /// <summary>Attempts to save an RTF document as PDF at the specified path and returns diagnostics instead of throwing.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this RtfDocument document, string path, RtfPdfSaveOptions? options = null) {
        try {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            return document.ToPdfDocument(options).TrySave(path);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>Attempts to save an RTF document as PDF to a stream and returns diagnostics instead of throwing.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this RtfDocument document, Stream stream, RtfPdfSaveOptions? options = null) {
        try {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            PdfCore.PdfSaveResult result = document.ToPdfDocument(options).TrySave(stream);
            return result;
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    /// <summary>Saves an RTF document as PDF and returns the generated bytes asynchronously.</summary>
    public static async Task<byte[]> SaveAsPdfAsync(this RtfDocument document, RtfPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        cancellationToken.ThrowIfCancellationRequested();
        using (MemoryStream stream = new MemoryStream()) {
            await document.ToPdfDocument(options).SaveAsync(stream, cancellationToken).ConfigureAwait(false);
            return stream.ToArray();
        }
    }

    /// <summary>Saves an RTF document as PDF at the specified path asynchronously.</summary>
    public static Task SaveAsPdfAsync(this RtfDocument document, string path, RtfPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        return document.ToPdfDocument(options).SaveAsync(path, cancellationToken);
    }

    /// <summary>Saves an RTF document as PDF to a writable stream asynchronously.</summary>
    public static async Task SaveAsPdfAsync(this RtfDocument document, Stream stream, RtfPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        await document.ToPdfDocument(options).SaveAsync(stream, cancellationToken).ConfigureAwait(false);
    }
}
