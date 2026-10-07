using System.ComponentModel;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using OfficeIMO.Tool.Agent;

namespace OfficeIMO.Tool.Mcp;

internal sealed partial class OfficeImoMcpTools {
    [McpServerTool(Name = "officeimo_pdf_print_plan", Title = "Plan PDF sheets without printing", ReadOnly = true,
        Destructive = false, Idempotent = true, OpenWorld = false, UseStructuredContent = true, OutputSchemaType = typeof(AgentPdfPrintPlanResult))]
    [Description("Return a bounded sheet-plan summary through the reusable print planner. Enforces PDF printing permissions. No printer queue or device is contacted.")]
    public Task<CallToolResult> PdfPrintPlanAsync(
        [Description("Local source PDF within allowed roots.")] string path,
        [Description("Optional document-relative page selection, at most10000 pages including repeats.")] string? pages = null,
        [Description("A4, Letter, Legal, or A3.")] string paper = "A4",
        [Description("auto, portrait, or landscape.")] string orientation = "auto",
        [Description("Source pages per sheet, 1, 2, or 4.")] int pagesPerSheet = 1,
        [Description("fit, actual, or fill.")] string scale = "fit",
        [Description("Uniform paper margin in points.")] double margin = 18,
        [Description("Name of the environment variable containing a source password.")] string? passwordEnvironmentVariable = null,
        [Description("Maximum serialized result characters, 512-64000.")] int maxOutputCharacters = 2000,
        CancellationToken cancellationToken = default) => ExecuteSafePdfAsync(() => _service.PdfPrintPlanAsync(path, pages, paper,
            orientation, pagesPerSheet, scale, margin, passwordEnvironmentVariable, maxOutputCharacters, cancellationToken),
            result => "Planned " + result.SheetCount + " sheet(s); no printing was requested.");

    [McpServerTool(Name = "officeimo_pdf_ocr_providers", Title = "List configured PDF OCR providers", ReadOnly = true,
        Destructive = false, Idempotent = true, OpenWorld = false, UseStructuredContent = true, OutputSchemaType = typeof(AgentPdfOcrProvidersResult))]
    [Description("List the OCR provider ids explicitly registered by the trusted server host. No ambient plugin scanning or installation occurs.")]
    public CallToolResult PdfOcrProviders([Description("Maximum serialized result characters, 512-64000.")] int maxOutputCharacters = 2000) =>
        Execute(() => _service.PdfOcrProviders(maxOutputCharacters), result => "Configured " + result.ProviderCount + " OCR provider(s).");

    [McpServerTool(Name = "officeimo_pdf", Title = "Create a separate PDF copy", ReadOnly = false,
        Destructive = true, Idempotent = true, OpenWorld = false, UseStructuredContent = true, OutputSchemaType = typeof(AgentPdfWorkflowResult))]
    [Description("Extract ordered pages, owner-authorized decrypt, fully raster flatten, optimize, or sanitize a PDF through bounded first-party workflows. Requires a separate explicit destination; signatures and mutation permissions retain the engine's policy. Raster output requires acknowledgement of lost native/interactive data. Returns artifact metadata and diagnostic codes, never text.")]
    public Task<CallToolResult> PdfAsync(
        [Description("Local source PDF within allowed roots.")] string path,
        [Description("Separate destination .pdf within allowed roots.")] string outputPath,
        [Description("extract, decrypt, flatten, optimize, or sanitize.")] string operation,
        [Description("Ordered document-relative selection, for example last,1-3 or odd. Required for extract; optional for flatten.")] string? pages = null,
        [Description("Environment-variable name containing the source password. Decrypt requires the owner password; passwords never enter tool arguments/results.")] string? passwordEnvironmentVariable = null,
        [Description("Explicit permission to replace an existing output copy.")] bool overwrite = false,
        [Description("Required for flatten: rendered appearances only; text, forms, links, signatures and attachments are omitted.")] bool acknowledgeRasterOutput = false,
        [Description("Maximum selected pages, 1-10000. Default100.")] int maximumPages = 100,
        [Description("Flatten sampling DPI, 36-600.")] double dpi = 150,
        [Description("Maximum input bytes, 1-268435456.")] long maximumInputBytes = 268435456,
        [Description("Maximum output bytes, 1-536870912.")] long maximumOutputBytes = 536870912,
        [Description("Maximum serialized result characters, 512-64000.")] int maxOutputCharacters = 4000,
        CancellationToken cancellationToken = default) => ExecutePdfAsync(() => _service.PdfAsync(path, outputPath, operation,
            Settings(pages, passwordEnvironmentVariable, overwrite, maximumPages, dpi, maximumInputBytes, maximumOutputBytes, acknowledgeRasterOutput),
            maxOutputCharacters, cancellationToken));

    [McpServerTool(Name = "officeimo_pdf_split", Title = "Split a PDF into separate parts", ReadOnly = false,
        Destructive = true, Idempotent = true, OpenWorld = false, UseStructuredContent = true, OutputSchemaType = typeof(AgentPdfWorkflowResult))]
    [Description("Split consecutive pages into a separate explicit output folder, reopen every part, and publish the folder as a unit. A destination containing the source is rejected. A bounded artifact sample may omit files; artifactCount describes the complete output.")]
    public Task<CallToolResult> SplitPdfAsync(
        [Description("Local source PDF within allowed roots.")] string path,
        [Description("Separate output folder within allowed roots.")] string outputDirectory,
        [Description("Consecutive pages per part, 1-10000.")] int pagesPerDocument = 1,
        [Description("Maximum parts, 1-10000.")] int maximumParts = 100,
        [Description("Name of the environment variable containing a source password.")] string? passwordEnvironmentVariable = null,
        [Description("Explicit permission to replace an existing destination folder and its files.")] bool overwrite = false,
        [Description("Maximum input bytes, 1-268435456.")] long maximumInputBytes = 268435456,
        [Description("Maximum aggregate output bytes, 1-536870912.")] long maximumOutputBytes = 536870912,
        [Description("Maximum serialized result characters, 512-64000.")] int maxOutputCharacters = 4000,
        CancellationToken cancellationToken = default) => ExecutePdfAsync(() => _service.SplitPdfAsync(path, outputDirectory, pagesPerDocument,
            Settings(null, passwordEnvironmentVariable, overwrite, maximumParts, 150, maximumInputBytes, maximumOutputBytes), maxOutputCharacters, cancellationToken));

    [McpServerTool(Name = "officeimo_pdf_ocr", Title = "Create a searchable PDF copy", ReadOnly = false,
        Destructive = true, Idempotent = false, OpenWorld = true, UseStructuredContent = true, OutputSchemaType = typeof(AgentPdfWorkflowResult))]
    [Description("Add searchable text through a provider explicitly registered by the server host. Provider assemblies, executables, model paths and options are trusted startup configuration and cannot be supplied through this tool. Preserves page count and returns bounded metadata; confidence is recognition evidence, not accuracy proof.")]
    public Task<CallToolResult> SearchablePdfAsync(
        [Description("Local source PDF within allowed roots.")] string path,
        [Description("Separate destination .pdf within allowed roots.")] string outputPath,
        [Description("Registered provider id, for example tesseract-cli.")] string providerId,
        [Description("Optional document-relative recognition page selection.")] string? pages = null,
        [Description("Provider language expression, for example eng or eng+pol.")] string? language = null,
        [Description("Minimum accepted provider confidence, 0-1.")] double minimumConfidence = 0.5,
        [Description("Name of the environment variable containing a source password.")] string? passwordEnvironmentVariable = null,
        [Description("Explicit permission to replace an existing output copy.")] bool overwrite = false,
        [Description("Maximum recognized pages, 1-10000.")] int maximumPages = 100,
        [Description("Recognition raster DPI, 36-600.")] double dpi = 150,
        [Description("Maximum input bytes, 1-268435456.")] long maximumInputBytes = 268435456,
        [Description("Maximum output bytes, 1-536870912.")] long maximumOutputBytes = 536870912,
        [Description("Maximum serialized result characters, 512-64000.")] int maxOutputCharacters = 4000,
        CancellationToken cancellationToken = default) => ExecutePdfAsync(() => _service.SearchablePdfAsync(path, outputPath, providerId,
            Settings(pages, passwordEnvironmentVariable, overwrite, maximumPages, dpi, maximumInputBytes, maximumOutputBytes),
            language, minimumConfidence, maxOutputCharacters, cancellationToken));

    [McpServerTool(Name = "officeimo_pdf_export_pages", Title = "Render selected PDF pages", ReadOnly = false,
        Destructive = true, Idempotent = true, OpenWorld = false, UseStructuredContent = true, OutputSchemaType = typeof(AgentPdfWorkflowResult))]
    [Description("Render selected pages into a separate explicit folder through the reusable image-export workflow. Outputs are decoded and validated before publication. Formats: png, jpeg, webp, tiff, svg.")]
    public Task<CallToolResult> ExportPdfPagesAsync(
        [Description("Local source PDF within allowed roots.")] string path,
        [Description("Separate output folder within allowed roots.")] string outputDirectory,
        [Description("png, jpeg, webp, tiff, or svg.")] string format = "png",
        [Description("Optional document-relative page selection.")] string? pages = null,
        [Description("Name of the environment variable containing a source password.")] string? passwordEnvironmentVariable = null,
        [Description("Explicit permission to replace an existing output folder and its files.")] bool overwrite = false,
        [Description("Maximum exported pages, 1-10000.")] int maximumPages = 100,
        [Description("Sampling DPI, 36-600.")] double dpi = 150,
        [Description("Maximum serialized result characters, 512-64000.")] int maxOutputCharacters = 4000,
        CancellationToken cancellationToken = default) => ExecutePdfAsync(() => _service.ExportPdfPagesAsync(path, outputDirectory, format,
            Settings(pages, passwordEnvironmentVariable, overwrite, maximumPages, dpi, 268435456, 536870912), maxOutputCharacters, cancellationToken));

    [McpServerTool(Name = "officeimo_pdf_assemble", Title = "Assemble PDF and image files", ReadOnly = false,
        Destructive = true, Idempotent = true, OpenWorld = false, UseStructuredContent = true, OutputSchemaType = typeof(AgentPdfWorkflowResult))]
    [Description("Assemble ordered explicit local PDF/raster-image files into one separate PDF. All source paths and the destination must be within allowed roots. Folder, archive and heterogeneous Office intake remain available through workflow assemble in the CLI.")]
    public Task<CallToolResult> AssemblePdfAsync(
        [Description("From 1 through 100 ordered local PDF/raster-image paths.")] string[] paths,
        [Description("Separate destination .pdf within allowed roots.")] string outputPath,
        [Description("Name of the environment variable containing a source-PDF password.")] string? passwordEnvironmentVariable = null,
        [Description("Explicit permission to replace an existing output copy.")] bool overwrite = false,
        [Description("Maximum aggregate input bytes, 1-268435456.")] long maximumInputBytes = 268435456,
        [Description("Maximum output bytes, 1-536870912.")] long maximumOutputBytes = 536870912,
        [Description("Maximum serialized result characters, 512-64000.")] int maxOutputCharacters = 4000,
        CancellationToken cancellationToken = default) => ExecutePdfAsync(() => _service.AssemblePdfAsync(paths, outputPath,
            Settings(null, passwordEnvironmentVariable, overwrite, 100, 150, maximumInputBytes, maximumOutputBytes), maxOutputCharacters, cancellationToken));

    private static PdfWorkflowSettings Settings(string? pages, string? password, bool overwrite, int maximumPages, double dpi,
        long inputBytes, long outputBytes, bool acknowledge = false) => new() {
        Pages = pages, PasswordEnvironmentVariable = password, Overwrite = overwrite, MaximumPages = maximumPages,
        Dpi = dpi, MaximumInputBytes = inputBytes, MaximumOutputBytes = outputBytes, AcknowledgeRasterOutput = acknowledge
    };

    private static Task<CallToolResult> ExecutePdfAsync(Func<Task<AgentPdfWorkflowResult>> action) =>
        ExecuteSafePdfAsync(action, result => "PDF " + result.Operation + ": " + result.Status + "; " + result.ArtifactCount + " artifact(s).", result => result.Succeeded);

    private static async Task<CallToolResult> ExecuteSafePdfAsync<T>(Func<Task<T>> action, Func<T, string> summary,
        Func<T, bool>? succeeded = null) where T : class {
        try {
            T result = await action().ConfigureAwait(false);
            CallToolResult response = Success(result, summary(result));
            response.IsError = succeeded is not null && !succeeded(result);
            return response;
        } catch (AgentUsageException exception) { return Error(exception.Message); }
        catch (UnauthorizedAccessException) { return Error("PDF access was refused. Check allowed roots, file permissions, and trusted provider configuration."); }
        catch (FileNotFoundException) { return Error("The selected local input or configured OCR provider assembly was not found."); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
            return Error("PDF operation failed (" + exception.GetType().Name + "). Check host configuration and diagnostic codes; no provider message or document text is returned.");
        }
    }
}
