using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Analyzes a file-to-file Excel conversion without creating or changing an artifact.</summary>
    public static ExcelDocumentConversionReport AnalyzeConversion(
        string sourcePath,
        string destinationPath,
        ExcelDocumentConversionOptions? options = null) =>
        AnalyzeConversionAsync(sourcePath, destinationPath, options).GetAwaiter().GetResult();

    /// <summary>Asynchronously analyzes a file-to-file Excel conversion without creating or changing an artifact.</summary>
    public static async Task<ExcelDocumentConversionReport> AnalyzeConversionAsync(
        string sourcePath,
        string destinationPath,
        ExcelDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options ??= new ExcelDocumentConversionOptions();
        OfficeFileConversion.Paths paths = OfficeFileConversion.ValidatePaths(
            sourcePath,
            destinationPath,
            SupportedExcelConversionExtensions,
            "Excel workbook");
        using ExcelDocument document = await LoadExcelConversionSourceAsync(paths.Source, options, cancellationToken)
            .ConfigureAwait(false);
        OfficeFormatDescriptor sourceDescriptor = document.SourceFormatDescriptor;
        OfficeFormatDescriptor destinationDescriptor = ExcelFormatCatalog.GetByExtension(paths.Destination);
        OfficeCompatibilityMode mode = GetCompatibilityMode(options);
        bool allowsLoss = AllowsLoss(options, mode);
        IReadOnlyList<ExcelConversionDiagnostic> diagnostics = CreateExcelConversionDiagnostics(
            document,
            paths.Source,
            sourceDescriptor,
            destinationDescriptor,
            options,
            mode,
            allowsLoss,
            out _,
            out _);

        return new ExcelDocumentConversionReport(
            paths.Source,
            paths.Destination,
            document.SourceFormat,
            GetExcelFormat(paths.Destination),
            sourceDescriptor,
            destinationDescriptor,
            diagnostics,
            mode,
            replacedExistingFile: false);
    }
}
