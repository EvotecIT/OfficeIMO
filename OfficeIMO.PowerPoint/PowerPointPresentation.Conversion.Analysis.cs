using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointPresentation {
    /// <summary>Analyzes a file-to-file PowerPoint conversion without creating or changing an artifact.</summary>
    public static PowerPointPresentationConversionReport AnalyzeConversion(
        string sourcePath,
        string destinationPath,
        PowerPointPresentationConversionOptions? options = null) =>
        AnalyzeConversionAsync(sourcePath, destinationPath, options).GetAwaiter().GetResult();

    /// <summary>Asynchronously analyzes a file-to-file PowerPoint conversion without creating or changing an artifact.</summary>
    public static async Task<PowerPointPresentationConversionReport> AnalyzeConversionAsync(
        string sourcePath,
        string destinationPath,
        PowerPointPresentationConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options ??= new PowerPointPresentationConversionOptions();
        OfficeFileConversion.Paths paths = OfficeFileConversion.ValidatePaths(
            sourcePath,
            destinationPath,
            SupportedPowerPointConversionExtensions,
            "PowerPoint presentation");
        using PowerPointPresentation presentation = await LoadAsync(
            paths.Source,
            CreateConversionLoadOptions(options.LoadOptions),
            cancellationToken).ConfigureAwait(false);
        OfficeFormatDescriptor sourceDescriptor = presentation.SourceFormatDescriptor;
        OfficeFormatDescriptor destinationDescriptor = PowerPointFormatCatalog.GetByExtension(paths.Destination);
        OfficeCompatibilityMode mode = GetCompatibilityMode(options);
        bool allowsLoss = AllowsLoss(options, mode);
        IReadOnlyList<PowerPointConversionDiagnostic> diagnostics = CreatePowerPointConversionDiagnostics(
            presentation,
            paths.Source,
            sourceDescriptor,
            destinationDescriptor,
            options,
            mode,
            allowsLoss,
            out _);

        return new PowerPointPresentationConversionReport(
            paths.Source,
            paths.Destination,
            presentation.SourceFormat,
            PowerPointPresentationLoadRouting.GetFormat(paths.Destination),
            sourceDescriptor,
            destinationDescriptor,
            mode,
            diagnostics,
            replacedExistingFile: false);
    }
}
