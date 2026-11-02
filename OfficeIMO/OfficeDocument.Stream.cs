using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

namespace OfficeIMO;

public static partial class OfficeDocument {
    /// <summary>Inspects bounded encoded content without filesystem access. A seekable source is read from the beginning and its position is restored; the caller retains ownership.</summary>
    public static OfficeDocumentInfo Inspect(Stream source, string sourceName, OfficeInspectOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (string.IsNullOrWhiteSpace(sourceName)) throw new ArgumentException("A source filename is required.", nameof(sourceName));
        OfficeInspectOptions settings = options ?? new();
        if (settings.Formats == null || ((settings.IncludeContent || settings.IncludeSummary) && settings.Reader == null)) throw new ArgumentException("Reader and format settings are required.", nameof(options));
        if (settings.MaximumInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options));
        if (settings.VbaSignatures != null || settings.WordMacroSignatures != null)
            throw new NotSupportedException("VBA and Word macro signature inspection require a source path.");
        long maximumBytes = settings.MaximumInputBytes;
        if ((settings.IncludeContent || settings.IncludeSummary) && settings.Reader.MaxInputBytes is long readerMaximum) maximumBytes = Math.Min(maximumBytes, readerMaximum);
        if (settings.ContentSafety != null) maximumBytes = Math.Min(maximumBytes, settings.ContentSafety.MaxInputBytes);
        if (settings.Provenance != null) maximumBytes = Math.Min(maximumBytes, settings.Provenance.MaxAssetBytes);
        if (settings.PackageSignatures != null) maximumBytes = Math.Min(maximumBytes, settings.PackageSignatures.MaxPackageBytes);
        if (settings.PdfPreflight != null) maximumBytes = Math.Min(maximumBytes, settings.PdfPreflight.Limits.MaxInputBytes);
        byte[] bytes = OfficeStreamReader.ReadAllBytes(source, cancellationToken, maximumBytes);
        string extension = Path.GetExtension(sourceName).ToLowerInvariant();
        var result = new OfficeDocumentInfo(sourceName);
        if (settings.PdfPreflight != null) result.PdfPreflight = PdfDocument.Preflight(bytes, settings.PdfPreflight, cancellationToken);
        if (settings.ContentSafety != null) result.ContentSafety = extension switch {
            ".docx" or ".docm" or ".dotx" or ".dotm" => WordDocument.InspectContentSafety(bytes, settings.ContentSafety, cancellationToken),
            ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xlsb" or ".xls" or ".xlam" => ExcelDocument.InspectContentSafety(bytes, sourceName, settings.ContentSafety),
            ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" => PowerPointDocument.InspectContentSafety(bytes, sourceName, settings.ContentSafety),
            ".pdf" => PdfDocument.InspectContentSafety(bytes, settings.ContentSafety, settings.Formats.Pdf?.ReadOptions),
            _ => throw new NotSupportedException("Concealed-content inspection is not supported for this source format.")
        };
        cancellationToken.ThrowIfCancellationRequested();
        if (settings.Provenance != null) {
            OfficeProvenanceOptions provenance = settings.Provenance.Snapshot(cancellationToken);
            result.Provenance = extension switch {
                ".docx" or ".docm" or ".dotx" or ".dotm" => WordDocument.InspectProvenance(bytes, sourceName, provenance),
                ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xlsb" or ".xlam" => ExcelDocument.InspectProvenance(bytes, sourceName, provenance),
                ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" => PowerPointDocument.InspectProvenance(bytes, sourceName, provenance),
                ".pdf" => PdfProvenance.Inspect(bytes, provenance, settings.Formats.Pdf?.ReadOptions),
                _ => throw new NotSupportedException("Provenance inspection is not supported for this source format.")
            };
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (settings.PackageSignatures != null) result.PackageSignatures = OfficePackageSignatureService.Inspect(bytes, settings.PackageSignatures);
        if (settings.IncludeSummary) {
            using var input = new MemoryStream(bytes, writable: false);
            InspectSummary(result, input, settings, cancellationToken);
        }
        if (settings.IncludeContent && !settings.IncludeSummary) result.Content = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers(settings.Formats).Build()
            .ReadDocument(bytes, sourceName, settings.Reader, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }
}
