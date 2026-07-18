using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using System.Threading;

namespace OfficeIMO.Word.OpenDocument;

/// <summary>Thin ODT image-export bridge over the Word visual renderer.</summary>
public static class WordOpenDocumentImageExportExtensions {
    /// <summary>Converts an ODT document to Word semantics and exports one selected page.</summary>
    public static OfficeImageExportResult ExportImage(
        this OdtDocument source,
        OfficeImageExportFormat format,
        WordImageExportOptions? imageOptions = null,
        WordOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        WordImageExportOptions effective =
            imageOptions?.Clone() ?? new WordImageExportOptions();
        OdfConversionResult<WordDocument> conversion =
            source.ToWordDocumentResult(conversionOptions);
        using (conversion.Value) {
            OfficeImageExportResult image =
                conversion.Value.ExportImage(format, effective, cancellationToken);
            return effective.EnsureAccepted(
                OdfImageExportDiagnostics.Attach(image, conversion.Report));
        }
    }

    /// <summary>Converts an ODT document to Word semantics and exports selected estimated pages.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this OdtDocument source,
        OfficeImageExportFormat format,
        WordImageExportOptions? imageOptions = null,
        WordOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        source.ExportImages(
            format,
            results.Add,
            imageOptions,
            conversionOptions,
            cancellationToken);
        return results.AsReadOnly();
    }

    /// <summary>Streams selected ODT page images without retaining earlier payloads.</summary>
    public static void ExportImages(
        this OdtDocument source,
        OfficeImageExportFormat format,
        OfficeImageExportConsumer consumer,
        WordImageExportOptions? imageOptions = null,
        WordOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        WordImageExportOptions effective =
            imageOptions?.Clone() ?? new WordImageExportOptions();
        OdfConversionResult<WordDocument> conversion =
            source.ToWordDocumentResult(conversionOptions);
        using (conversion.Value) {
            conversion.Value.ExportImages(
                format,
                result => consumer(effective.EnsureAccepted(
                    OdfImageExportDiagnostics.Attach(result, conversion.Report))),
                effective,
                cancellationToken);
        }
    }
}
