using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;

namespace OfficeIMO.Word.OpenDocument;

/// <summary>Thin ODT image-export bridge over the Word visual renderer.</summary>
public static partial class WordOpenDocumentImageExportExtensions {
    /// <summary>Converts an ODT document to Word semantics and exports one selected page.</summary>
    public static OfficeImageExportResult ToImage(
        this OdtDocument source,
        OfficeImageExportFormat format,
        WordImageExportOptions? imageOptions = null,
        WordOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(source);
        WordImageExportOptions effective =
            imageOptions?.Clone() ?? new WordImageExportOptions();
        return OfficeImageExportExecutionScope.Run(effective, cancellationToken, token => {
            OdfConversionResult<WordDocument> conversion =
                source.ConvertToWordWithEvidence(conversionOptions);
            using (conversion.Value) {
                OfficeImageExportResult image =
                    conversion.Value.ToImage(format, effective, token);
                return effective.EnsureAccepted(
                    OdfImageExportDiagnostics.Attach(image, conversion.Report));
            }
        });
    }

    /// <summary>Converts an ODT document to Word semantics and exports selected estimated pages.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(
        this OdtDocument source,
        OfficeImageExportFormat format,
        WordImageExportOptions? imageOptions = null,
        WordOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        source.ToImages(
            format,
            results.Add,
            imageOptions,
            conversionOptions,
            cancellationToken);
        return results.AsReadOnly();
    }

    /// <summary>Streams selected ODT page images without retaining earlier payloads.</summary>
    public static void ToImages(
        this OdtDocument source,
        OfficeImageExportFormat format,
        OfficeImageExportConsumer consumer,
        WordImageExportOptions? imageOptions = null,
        WordOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(consumer);
        WordImageExportOptions effective =
            imageOptions?.Clone() ?? new WordImageExportOptions();
        OfficeImageExportExecutionScope.Run(effective, cancellationToken, token => {
            OdfConversionResult<WordDocument> conversion =
                source.ConvertToWordWithEvidence(conversionOptions);
            using (conversion.Value) {
                conversion.Value.ToImages(
                    format,
                    result => consumer(effective.EnsureAccepted(
                        OdfImageExportDiagnostics.Attach(result, conversion.Report))),
                    effective,
                    token);
            }
            return 0;
        });
    }
}
