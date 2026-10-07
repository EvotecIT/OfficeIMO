using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    internal async Task<AgentPdfPrintPlanResult> PdfPrintPlanAsync(string path, string? pages, string paper,
        string orientation, int pagesPerSheet, string scale, double margin, string? passwordEnvironmentVariable,
        int maxOutputCharacters, CancellationToken cancellationToken) {
        maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        var settings = new PdfWorkflowSettings { Pages = pages, PasswordEnvironmentVariable = passwordEnvironmentVariable };
        settings.Validate();
        string input = ResolvePdfInput(path);
        await using Stream stream = await PdfInput(input).OpenRead(cancellationToken).ConfigureAwait(false);
        PdfDocument document = await PdfDocument.LoadAsync(stream, new PdfLoadOptions {
            Password = settings.Password(), Limits = new PdfReadLimits { MaxInputBytes = settings.MaximumInputBytes }
        }, cancellationToken).ConfigureAwait(false);
        var selector = settings.Selector();
        if (selector is not null) _ = selector.Resolve(document.Inspect().PageCount, 10_000);
        else if (document.Inspect().PageCount > 10_000) throw new AgentUsageException("Print-plan selection is limited to 10000 pages.");
        PdfPrintPlan plan = PdfPrintPlanner.Create(document, new PdfPrintPlanRequest {
            InputPath = input, Pages = pages, PagesPerSheet = pagesPerSheet, Margin = margin,
            PaperSize = paper.ToLowerInvariant() switch {
                "a4" => PageSizes.A4, "letter" => PageSizes.Letter, "legal" => PageSizes.Legal, "a3" => PageSizes.A3,
                _ => throw new AgentUsageException("paper must be A4, Letter, Legal, or A3.")
            },
            Orientation = orientation.ToLowerInvariant() switch {
                "auto" => PdfPrintOrientation.Automatic, "portrait" => PdfPrintOrientation.Portrait, "landscape" => PdfPrintOrientation.Landscape,
                _ => throw new AgentUsageException("orientation must be auto, portrait, or landscape.")
            },
            ScaleMode = scale.ToLowerInvariant() switch {
                "fit" => PdfPrintScaleMode.Fit, "actual" => PdfPrintScaleMode.ActualSize, "fill" => PdfPrintScaleMode.Fill,
                _ => throw new AgentUsageException("scale must be fit, actual, or fill.")
            }
        }, cancellationToken);
        var sampled = plan.SelectedPages.Take(25).ToList();
        var result = new AgentPdfPrintPlanResult {
            SourcePageCount = plan.SourcePageCount, SelectedPageCount = plan.SelectedPages.Count, SheetCount = plan.Sheets.Count,
            ClippedPlacementCount = plan.Sheets.Sum(sheet => sheet.Placements.Count(placement => placement.IsClipped)),
            SelectedPages = sampled, Truncated = sampled.Count < plan.SelectedPages.Count
        };
        while (AgentJson.Measure(result) > maxOutputCharacters) { sampled.RemoveAt(sampled.Count - 1); result.Truncated = true; }
        return result;
    }
}
