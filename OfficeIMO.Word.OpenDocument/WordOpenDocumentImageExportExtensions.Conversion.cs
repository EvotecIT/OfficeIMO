using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.OpenDocument;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentImageExportExtensions {
    /// <summary>Renders selected source surfaces using the encoding in the image options.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this OdtDocument document,
        WordImageExportOptions? imageOptions = null, WordOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) =>
        CreateImageOperation(document, imageOptions?.Clone() ?? new WordImageExportOptions(), conversionOptions).Export(cancellationToken);

    /// <summary>Renders selected source surfaces asynchronously under one shared operation deadline and budget.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(this OdtDocument document,
        WordImageExportOptions? imageOptions = null, WordOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        var resolved = imageOptions?.Clone() ?? new WordImageExportOptions();
        return CreateImageOperation(document, resolved, conversionOptions).ExportAsync(cancellationToken);
    }

    /// <summary>Saves selected source surfaces through the shared atomic file and fidelity policies.</summary>
    public static OfficeConversionReport SaveAsImages(this OdtDocument document, string folderPath,
        WordImageExportOptions? imageOptions = null, WordOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        var resolved = imageOptions?.Clone() ?? new WordImageExportOptions();
        var saved = CreateImageOperation(document, resolved, conversionOptions).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("ODT", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Saves selected source surfaces asynchronously without retaining earlier encoded payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this OdtDocument document, string folderPath,
        WordImageExportOptions? imageOptions = null, WordOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        var resolved = imageOptions?.Clone() ?? new WordImageExportOptions();
        var saved = await CreateImageOperation(document, resolved, conversionOptions)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("ODT", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<WordImageExportOptions> CreateImageOperation(
        OdtDocument document, WordImageExportOptions options, WordOpenDocumentConversionOptions? conversionOptions) {
        ArgumentNullException.ThrowIfNull(document);
        WordOpenDocumentConversionOptions? conversion = conversionOptions?.Clone();
        return new OfficeImageExportBatchOperation<WordImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (format, effective) => document.ToImages(format, effective, conversion),
            exportEach: (format, effective, consumer, token) => document.ToImages(format, consumer, effective, conversion, token));
    }
}
