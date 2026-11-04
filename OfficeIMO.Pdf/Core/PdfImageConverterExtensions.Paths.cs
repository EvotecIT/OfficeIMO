using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

/// <summary>Converts ordered image sources into PDF pages.</summary>
public static partial class PdfImageConverterExtensions {
    /// <summary>Converts each image to a PDF page in caller order.</summary>
    public static PdfDocument ToPdf(this IEnumerable<string> imagePaths, PdfImageImportOptions? options = null,
        CancellationToken cancellationToken = default) => imagePaths.ToPdf(options, out _, cancellationToken);

    /// <summary>Converts ordered images and returns completed rendering findings.</summary>
    public static PdfDocument ToPdf(this IEnumerable<string> imagePaths, PdfImageImportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfImageImportOptions operation = (options ?? new PdfImageImportOptions()).CloneAndValidate();
        return CompleteImageModel(() => PdfDocument.CreateFromImages(imagePaths, operation, cancellationToken),
            operation, cancellationToken, out report);
    }

    /// <summary>Returns the PDF model and completed rendering report together.</summary>
    public static async Task<OfficeConversionResult<PdfDocument>> ToPdfAsync(this IEnumerable<string> imagePaths,
        PdfImageImportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfImageImportOptions operation = (options ?? new PdfImageImportOptions()).CloneAndValidate();
        return await Task.Run(() => {
            PdfDocument value = imagePaths.ToPdf(operation, out OfficeConversionReport report, cancellationToken);
            return new OfficeConversionResult<PdfDocument>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes ordered image pages to a file and returns completed conversion findings.</summary>
    public static OfficeConversionReport SaveAsPdf(this IEnumerable<string> imagePaths, string path,
        PdfImageImportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        cancellationToken.ThrowIfCancellationRequested();
        PdfImageImportOptions operation = (options ?? new PdfImageImportOptions()).CloneAndValidate();
        return SaveImages(() => PdfDocument.CreateFromImages(imagePaths, operation, cancellationToken),
            (conversion, timer) => operation.FailOnLoss ? PdfConversionEvidence.SaveLossless(conversion, "Image", path, timer, cancellationToken) : conversion.Save(path, cancellationToken), cancellationToken);
    }

    /// <summary>Writes ordered image pages to a file and returns completed conversion findings.</summary>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this IEnumerable<string> imagePaths, string path,
        PdfImageImportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        cancellationToken.ThrowIfCancellationRequested();
        PdfImageImportOptions operation = (options ?? new PdfImageImportOptions()).CloneAndValidate();
        return await SaveImagesAsync(() => PdfDocument.CreateFromImages(imagePaths, operation, cancellationToken),
            (conversion, timer) => operation.FailOnLoss ? PdfConversionEvidence.SaveLosslessAsync(conversion, "Image", path, timer, cancellationToken) : conversion.SaveAsync(path, cancellationToken), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes ordered image pages to a caller-owned stream and returns completed conversion findings.</summary>
    public static OfficeConversionReport SaveAsPdf(this IEnumerable<string> imagePaths, Stream stream,
        PdfImageImportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        cancellationToken.ThrowIfCancellationRequested();
        PdfImageImportOptions operation = (options ?? new PdfImageImportOptions()).CloneAndValidate();
        return SaveImages(() => PdfDocument.CreateFromImages(imagePaths, operation, cancellationToken),
            (conversion, timer) => operation.FailOnLoss ? PdfConversionEvidence.SaveLossless(conversion, "Image", stream, timer, cancellationToken) : conversion.Save(stream, cancellationToken), cancellationToken);
    }

    /// <summary>Writes ordered image pages to a caller-owned stream and returns completed conversion findings.</summary>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this IEnumerable<string> imagePaths, Stream stream,
        PdfImageImportOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        cancellationToken.ThrowIfCancellationRequested();
        PdfImageImportOptions operation = (options ?? new PdfImageImportOptions()).CloneAndValidate();
        return await SaveImagesAsync(() => PdfDocument.CreateFromImages(imagePaths, operation, cancellationToken),
            (conversion, timer) => operation.FailOnLoss ? PdfConversionEvidence.SaveLosslessAsync(conversion, "Image", stream, timer, cancellationToken) : conversion.SaveAsync(stream, cancellationToken), cancellationToken).ConfigureAwait(false);
    }
}
