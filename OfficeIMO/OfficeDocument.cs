using OfficeIMO.ContentSafety;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using OfficeIMO.Word;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

namespace OfficeIMO;

/// <summary>Inspects documents through the native format handlers included by the front-door package.</summary>
public static partial class OfficeDocument {
    /// <summary>Inspects a local file using default inspection options.</summary>
    public static OfficeDocumentInfo Inspect(string path) => Inspect(path, null, default);

    /// <summary>Inspects a local file using the supplied inspection options.</summary>
    public static OfficeDocumentInfo Inspect(string path, OfficeInspectOptions? options) => Inspect(path, options, default);

    /// <summary>Returns bounded content, native inspection metadata and explicit coverage findings without modifying the source.</summary>
    public static OfficeDocumentInfo Inspect(string path, OfficeInspectOptions? options,
        CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A document path is required.", nameof(path));
        OfficeInspectOptions settings = options ?? new();
        if (settings.Formats == null || ((settings.IncludeContent || settings.IncludeSummary) && settings.Reader == null)) throw new ArgumentException("Reader and format settings are required.", nameof(options));
        string source = Path.GetFullPath(path);
        cancellationToken.ThrowIfCancellationRequested();
        var result = new OfficeDocumentInfo(source);
        string extension = Path.GetExtension(source);
        if (settings.PdfPreflight != null) {
            using var input = File.OpenRead(source);
            byte[] bytes = OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(input, cancellationToken, settings.PdfPreflight.Limits.MaxInputBytes);
            result.PdfPreflight = PdfDocument.Preflight(bytes, settings.PdfPreflight, cancellationToken);
        }
        if (settings.ContentSafety != null) {
            result.ContentSafety = extension.ToLowerInvariant() switch {
                ".docx" or ".docm" or ".dotx" or ".dotm" => WordDocument.InspectContentSafety(source, settings.ContentSafety, cancellationToken),
                ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xlsb" or ".xls" or ".xlam" => ExcelDocument.InspectContentSafety(source, settings.ContentSafety),
                ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" => PowerPointDocument.InspectContentSafety(source, settings.ContentSafety),
                ".pdf" => PdfDocument.InspectContentSafety(source, settings.ContentSafety, settings.Formats.Pdf?.ReadOptions),
                _ => throw new NotSupportedException("Concealed-content inspection is not supported for this source format.")
            };
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (settings.Provenance != null) {
            OfficeProvenanceOptions provenance = settings.Provenance.Snapshot(cancellationToken);
            result.Provenance = extension.ToLowerInvariant() switch {
                ".docx" or ".docm" or ".dotx" or ".dotm" => WordDocument.InspectProvenance(source, provenance),
                ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xlsb" or ".xlam" => ExcelDocument.InspectProvenance(source, provenance),
                ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" => PowerPointDocument.InspectProvenance(source, provenance),
                ".pdf" => PdfProvenance.InspectFile(source, provenance, settings.Formats.Pdf?.ReadOptions),
                _ => throw new NotSupportedException("Provenance inspection is not supported for this source format.")
            };
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (settings.PackageSignatures != null) result.PackageSignatures = OfficePackageSignatureService.Inspect(source, settings.PackageSignatures);
        if (settings.VbaSignatures != null) result.VbaSignatures = OfficeVbaSignatureService.Inspect(source, settings.VbaSignatures);
        if (settings.WordMacroSignatures != null) result.WordMacroSignatures = settings.SecurityProvider == null
            ? WordDocument.InspectMacroProjectSignatures(source, settings.WordMacroSignatures)
            : WordDocument.InspectMacroProjectSignatures(source, settings.SecurityProvider, settings.WordMacroSignatures);
        cancellationToken.ThrowIfCancellationRequested();
        if (settings.IncludeSummary) {
            using var input = File.OpenRead(source);
            InspectSummary(result, input, settings, cancellationToken, sourceIsPath: true);
            result.Modified = File.GetLastWriteTimeUtc(source);
        }
        if (settings.IncludeContent && !settings.IncludeSummary) result.Content = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers(settings.Formats).Build()
            .ReadDocument(source, settings.Reader, cancellationToken);
        return result;
    }

}
