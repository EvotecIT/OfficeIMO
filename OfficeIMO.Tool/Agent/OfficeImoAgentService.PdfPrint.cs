using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    internal async Task<AgentPdfPrintPlanResult> PdfPrintPlanAsync(string path, string? pages, string paper,
        string orientation, int pagesPerSheet, string scale, double margin, string? passwordEnvironmentVariable,
        int maxOutputCharacters, CancellationToken cancellationToken, double customScalePercent = 100,
        string alignment = "center", string pageSubset = "all", string colorMode = "color",
        double? marginLeft = null, double? marginTop = null, double? marginRight = null, double? marginBottom = null) {
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
            MarginLeft = marginLeft, MarginTop = marginTop, MarginRight = marginRight, MarginBottom = marginBottom,
            CustomScalePercent = customScalePercent,
            Alignment = alignment.ToLowerInvariant() switch {
                "top-left" => PdfPrintAlignment.TopLeft, "top" => PdfPrintAlignment.Top, "top-right" => PdfPrintAlignment.TopRight,
                "left" => PdfPrintAlignment.Left, "center" => PdfPrintAlignment.Center, "right" => PdfPrintAlignment.Right,
                "bottom-left" => PdfPrintAlignment.BottomLeft, "bottom" => PdfPrintAlignment.Bottom, "bottom-right" => PdfPrintAlignment.BottomRight,
                _ => throw new AgentUsageException("alignment must be top-left, top, top-right, left, center, right, bottom-left, bottom, or bottom-right.")
            },
            PageSubset = pageSubset.ToLowerInvariant() switch {
                "all" => PdfPrintPageSubset.All, "odd" => PdfPrintPageSubset.Odd, "even" => PdfPrintPageSubset.Even,
                _ => throw new AgentUsageException("pageSubset must be all, odd, or even.")
            },
            ColorMode = colorMode.ToLowerInvariant() switch {
                "color" => PdfPrintColorMode.Color, "grayscale" => PdfPrintColorMode.Grayscale,
                _ => throw new AgentUsageException("colorMode must be color or grayscale.")
            },
            PaperSize = paper.ToLowerInvariant() switch {
                "a4" => PageSizes.A4, "letter" => PageSizes.Letter, "legal" => PageSizes.Legal, "a3" => PageSizes.A3,
                _ => throw new AgentUsageException("paper must be A4, Letter, Legal, or A3.")
            },
            Orientation = orientation.ToLowerInvariant() switch {
                "auto" => PdfPrintOrientation.Automatic, "portrait" => PdfPrintOrientation.Portrait, "landscape" => PdfPrintOrientation.Landscape,
                _ => throw new AgentUsageException("orientation must be auto, portrait, or landscape.")
            },
            ScaleMode = scale.ToLowerInvariant() switch {
                "fit" => PdfPrintScaleMode.Fit, "actual" => PdfPrintScaleMode.ActualSize, "fill" => PdfPrintScaleMode.Fill, "custom" => PdfPrintScaleMode.Custom,
                _ => throw new AgentUsageException("scale must be fit, actual, fill, or custom.")
            }
        }, cancellationToken);
        var sampled = plan.SelectedPages.Take(25).ToList();
        var result = new AgentPdfPrintPlanResult {
            SourcePageCount = plan.SourcePageCount, SelectedPageCount = plan.SelectedPages.Count, SheetCount = plan.Sheets.Count,
            ClippedPlacementCount = plan.Sheets.Sum(sheet => sheet.Placements.Count(placement => placement.IsClipped)),
            SelectedPages = sampled, Truncated = sampled.Count < plan.SelectedPages.Count,
            ColorMode = plan.ColorMode.ToString().ToLowerInvariant()
        };
        while (AgentJson.Measure(result) > maxOutputCharacters) { sampled.RemoveAt(sampled.Count - 1); result.Truncated = true; }
        return result;
    }
}
