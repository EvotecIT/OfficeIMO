using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Analyzes a file-to-file Word conversion without creating or changing an artifact.</summary>
    public static WordDocumentConversionReport AnalyzeConversion(
        string sourcePath,
        string destinationPath,
        WordDocumentConversionOptions? options = null) =>
        AnalyzeConversionAsync(sourcePath, destinationPath, options).GetAwaiter().GetResult();

    /// <summary>Asynchronously analyzes a file-to-file Word conversion without creating or changing an artifact.</summary>
    public static async Task<WordDocumentConversionReport> AnalyzeConversionAsync(
        string sourcePath,
        string destinationPath,
        WordDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options ??= new WordDocumentConversionOptions();
        OfficeFileConversion.Paths paths = OfficeFileConversion.ValidatePaths(
            sourcePath,
            destinationPath,
            SupportedWordConversionExtensions,
            "Word document");
        using WordDocument document = await LoadWordConversionSourceAsync(paths.Source, options, cancellationToken)
            .ConfigureAwait(false);
        OfficeFormatDescriptor sourceDescriptor = document.SourceFormatDescriptor;
        OfficeFormatDescriptor destinationDescriptor = WordFormatCatalog.GetByExtension(paths.Destination);
        OfficeCompatibilityMode mode = GetCompatibilityMode(options);
        bool allowsLoss = AllowsLoss(options, mode);
        IReadOnlyList<WordConversionDiagnostic> diagnostics = CreateWordConversionDiagnostics(
            document,
            paths.Source,
            sourceDescriptor,
            destinationDescriptor,
            options,
            mode,
            allowsLoss,
            out _,
            out _);

        return new WordDocumentConversionReport(
            paths.Source,
            paths.Destination,
            document.SourceFormat,
            GetWordFormat(paths.Destination),
            sourceDescriptor,
            destinationDescriptor,
            diagnostics,
            mode,
            replacedExistingFile: false);
    }
}
