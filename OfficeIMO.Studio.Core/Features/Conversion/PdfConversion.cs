using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Conversion;

/// <summary>A document format that converts to PDF: its route, its name and the file extensions it is read from.</summary>
internal sealed record PdfConversionSource(string Route, string Name, IReadOnlyList<string> Extensions);

/// <summary>A note about a conversion's fidelity, such as content the PDF shows differently; severity is "warning" or "error".</summary>
internal sealed record PdfConversionNote(string Severity, string Message);

/// <summary>A finished conversion: the route used, the PDF's size and the converter's notes.</summary>
internal sealed record PdfConversionResult(string Route, string Source, long OutputBytes, IReadOnlyList<PdfConversionNote> Notes);

/// <summary>
/// Converts documents to PDF through the same workflow routes as OfficeIMO Studio's Convert window: Word, Excel,
/// PowerPoint, RTF, text, HTML, Markdown, XPS, Visio, Draw and Publisher files. PDFs convert to Word, Excel,
/// PowerPoint and HTML the same way.
/// </summary>
internal static class PdfConversion {
    /// <summary>The formats that convert to PDF, one per route.</summary>
    internal static IReadOnlyList<PdfConversionSource> Sources { get; } = OfficeWorkflowCatalog.ExecutableRoutes
        .Where(static route => string.Equals(route.Target, "PDF", StringComparison.OrdinalIgnoreCase))
        .Select(static route => new PdfConversionSource(route.Id, route.Source, route.SourceExtensions.ToArray()))
        .ToArray();

    /// <summary>
    /// The route for a file: the one whose first extension is the file's, as text files are plain text before they are
    /// HTML or Markdown, then any that lists it.
    /// </summary>
    internal static PdfConversionSource? RouteFor(string path) {
        string extension = Path.GetExtension(path);
        if (string.IsNullOrEmpty(extension)) return null;
        return Sources.FirstOrDefault(source => string.Equals(source.Extensions.FirstOrDefault(), extension, StringComparison.OrdinalIgnoreCase))
            ?? Sources.FirstOrDefault(source => source.Extensions.Contains(extension, StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>Converts a local file to a PDF at <paramref name="outputPath"/>, replacing what is there.</summary>
    internal static async Task<PdfConversionResult> ConvertAsync(string inputPath, string outputPath, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(inputPath) || !Path.IsPathFullyQualified(inputPath))
            throw new ArgumentException("Give the document to convert as an absolute path.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath) || !Path.IsPathFullyQualified(outputPath))
            throw new ArgumentException("Give the PDF an absolute destination.", nameof(outputPath));
        if (!File.Exists(inputPath)) throw new FileNotFoundException("The document to convert was not found.", inputPath);
        PdfConversionSource route = RouteFor(inputPath)
            ?? throw new NotSupportedException($"“{Path.GetFileName(inputPath)}” is not a format OfficeIMO converts to PDF.");
        return await RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = inputPath,
            OutputPath = outputPath,
            ConversionRouteId = route.Route,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
        }, route.Name, $"“{Path.GetFileName(inputPath)}” couldn’t be converted", cancellationToken).ConfigureAwait(false);
    }

    /// <summary>The formats a PDF converts to, by the route that converts it: Word, Excel, PowerPoint and HTML.</summary>
    internal static IReadOnlyDictionary<string, string> PdfTargets { get; } = new Dictionary<string, string>(StringComparer.Ordinal) {
        ["word"] = "pdf-docx", ["excel"] = "pdf-xlsx", ["powerpoint"] = "pdf-pptx", ["html"] = "pdf-html"
    };

    /// <summary>
    /// Converts a PDF, such as the open revision of a document, to Word, Excel, PowerPoint or HTML at
    /// <paramref name="outputPath"/>, replacing what is there. A password opens an encrypted PDF.
    /// </summary>
    internal static async Task<PdfConversionResult> ConvertPdfAsync(byte[] pdf, string? password, string target, string outputPath,
        CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(pdf);
        if (!PdfTargets.TryGetValue(target, out string? routeId) || OfficeWorkflowCatalog.FindExecutable(routeId) is null)
            throw new ArgumentException($"“{target}” isn't a format a PDF converts to.", nameof(target));
        if (string.IsNullOrWhiteSpace(outputPath) || !Path.IsPathFullyQualified(outputPath))
            throw new ArgumentException("Give the converted document an absolute destination.", nameof(outputPath));
        // The runner reads a file; the revision goes to a private temporary folder that is removed afterwards.
        string folder = Directory.CreateTempSubdirectory("officeimo-convert-").FullName;
        try {
            string inputPath = Path.Combine(folder, "document.pdf");
            await File.WriteAllBytesAsync(inputPath, pdf, cancellationToken).ConfigureAwait(false);
            return await RunAsync(new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert,
                InputPath = inputPath,
                OutputPath = outputPath,
                ConversionRouteId = routeId,
                PdfPassword = password,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            }, "PDF", "The PDF couldn’t be converted", cancellationToken).ConfigureAwait(false);
        } finally {
            try { Directory.Delete(folder, recursive: true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }
    }

    private static async Task<PdfConversionResult> RunAsync(OfficeWorkflowRequest request, string source, string failure,
        CancellationToken cancellationToken) {
        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(request, progress: null, cancellationToken).ConfigureAwait(false);
        switch (result.Status) {
            case OfficeWorkflowStatus.Completed:
                return new PdfConversionResult(request.ConversionRouteId!, source, result.OutputBytes, result.Diagnostics
                    .Where(static diagnostic => diagnostic.Severity != OfficeWorkflowDiagnosticSeverity.Information)
                    .Select(static diagnostic => new PdfConversionNote(diagnostic.Severity == OfficeWorkflowDiagnosticSeverity.Error ? "error" : "warning",
                        diagnostic.Message)).ToArray());
            case OfficeWorkflowStatus.Cancelled:
                throw new OperationCanceledException(cancellationToken);
            default:
                string detail = result.Diagnostics.FirstOrDefault(static diagnostic => diagnostic.Severity == OfficeWorkflowDiagnosticSeverity.Error)?.Message
                    ?? result.Summary;
                throw new InvalidOperationException($"{failure}: {detail}");
        }
    }
}
