using OfficeIMO.Drawing;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    internal async Task<AgentPdfWorkflowResult> ExportPdfPagesAsync(string path, string outputDirectory,
        string format, PdfWorkflowSettings settings, int maxOutputCharacters = 4000, CancellationToken cancellationToken = default) {
        settings.Validate(); maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        OfficeImageExportFormat imageFormat = format.ToLowerInvariant() switch {
            "png" => OfficeImageExportFormat.Png, "jpeg" or "jpg" => OfficeImageExportFormat.Jpeg,
            "webp" => OfficeImageExportFormat.Webp, "tiff" or "tif" => OfficeImageExportFormat.Tiff,
            "svg" => OfficeImageExportFormat.Svg,
            _ => throw new AgentUsageException("format must be png, jpeg, webp, tiff, or svg.")
        };
        string input = ResolvePdfInput(path);
        string destination = PreparePdfOutput(outputDirectory, [input], true, settings.Overwrite, "export-pages", maxOutputCharacters);
        PdfPageImageExportResult result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(new PdfPageImageExportRequest {
            InputPath = input, InputStream = PdfInput(input), OutputDirectory = destination, Pages = settings.Pages,
            Format = imageFormat, TargetDpi = settings.Dpi, MaximumPages = settings.MaximumPages,
            ConflictPolicy = settings.ConflictPolicy, PdfPassword = settings.Password(), Limits = settings.Limits(),
            PublicationGuard = new PdfRootPublicationGuard(_pathPolicy, [input])
        }, cancellationToken: cancellationToken).ConfigureAwait(false);
        return PdfResult("export-pages", result.Status, result.FailureKind, result.OutputDirectory, result.OutputBytes,
            result.Files.Select(file => new AgentPdfArtifact { Path = file.Path, SizeBytes = file.SizeBytes, FirstSourcePage = file.PageNumber, PageCount = 1 }).ToArray(),
            result.Diagnostics, maxOutputCharacters);
    }

    internal async Task<AgentPdfWorkflowResult> AssemblePdfAsync(IReadOnlyList<string> paths, string outputPath,
        PdfWorkflowSettings settings, int maxOutputCharacters = 4000, CancellationToken cancellationToken = default) {
        settings.Validate(); maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        if (paths is null || paths.Count is < 1 or > 100) throw new AgentUsageException("Assembly requires from 1 through 100 local source files.");
        string[] sources = paths.Select(_pathPolicy.ResolveInput).ToArray();
        if (sources.Any(source => !File.Exists(source))) throw new AgentUsageException("Agent assembly accepts explicit local files. Folder and ZIP intake are available through workflow assemble.");
        string[] accepted = [".pdf", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp", ".gif"];
        if (sources.Any(source => !accepted.Contains(Path.GetExtension(source), StringComparer.OrdinalIgnoreCase)))
            throw new AgentUsageException("Agent assembly accepts PDF and raster image files. Heterogeneous Office, folder, and ZIP intake are available through workflow assemble.");
        string destination = PreparePdfOutput(outputPath, sources, false, settings.Overwrite, "assemble", maxOutputCharacters);
        PdfAssemblyResult result = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = sources, SourceStreams = sources.Distinct(StringComparer.Ordinal).ToDictionary(source => source, PdfInput, StringComparer.Ordinal),
            OutputPath = destination, PdfPassword = settings.Password(), ConflictPolicy = settings.ConflictPolicy,
            Limits = settings.Limits(), Options = new PdfAssemblyOptions { MaximumSourceCount = 100 },
            PublicationGuard = new PdfRootPublicationGuard(_pathPolicy, sources)
        }, cancellationToken: cancellationToken).ConfigureAwait(false);
        return PdfResult("assemble", result.Status, result.FailureKind, result.OutputPath, result.OutputBytes,
            result.OutputPath is null ? [] : [new AgentPdfArtifact { Path = result.OutputPath, SizeBytes = result.OutputBytes, PageCount = result.PageCount }],
            result.Diagnostics, maxOutputCharacters);
    }
}
