using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
    /// <summary>Analyzes a file-to-file PowerPoint conversion without creating or changing an artifact.</summary>
    internal static PowerPointDocumentConversionReport AnalyzeConversion(
        string sourcePath,
        string destinationPath,
        PowerPointDocumentConversionOptions? options = null) =>
        AnalyzeConversionAsync(sourcePath, destinationPath, options).GetAwaiter().GetResult();

    /// <summary>Asynchronously analyzes a file-to-file PowerPoint conversion without creating or changing an artifact.</summary>
    internal static async Task<PowerPointDocumentConversionReport> AnalyzeConversionAsync(
        string sourcePath,
        string destinationPath,
        PowerPointDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options ??= new PowerPointDocumentConversionOptions();
        OfficeFileConversion.Paths paths = OfficeFileConversion.ValidatePaths(
            sourcePath,
            destinationPath,
            SupportedPowerPointConversionExtensions,
            "PowerPoint presentation");
        using PowerPointDocument presentation = await LoadAsync(
            paths.Source,
            CreateConversionLoadOptions(options.LoadOptions),
            cancellationToken).ConfigureAwait(false);
        OfficeFormatDescriptor sourceDescriptor = presentation.SourceFormatDescriptor;
        OfficeFormatDescriptor destinationDescriptor = PowerPointFormatCatalog.GetByExtension(paths.Destination);
        OfficeCompatibilityMode mode = GetCompatibilityMode(options);
        bool allowsLoss = AllowsLoss(options, mode);
        IReadOnlyList<OfficeConversionDiagnostic> diagnostics = CreateOfficeConversionDiagnostics(
            presentation,
            paths.Source,
            sourceDescriptor,
            destinationDescriptor,
            options,
            mode,
            allowsLoss,
            out _);

        return new PowerPointDocumentConversionReport(
            paths.Source,
            paths.Destination,
            presentation.SourceFormat,
            PowerPointDocumentLoadRouting.GetFormat(paths.Destination),
            sourceDescriptor,
            destinationDescriptor,
            mode,
            diagnostics,
            replacedExistingFile: false);
    }
}
