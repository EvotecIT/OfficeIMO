using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

/// <summary>
/// Capability-gallery helpers for Excel HTML exports.
/// </summary>
public static class ExcelHtmlCapabilityGalleryExtensions {
    /// <summary>
    /// Saves semantic and visual Excel HTML artifacts plus shared manifest files.
    /// </summary>
    /// <param name="workbook">Workbook to export.</param>
    /// <param name="directoryPath">Directory where artifacts will be written.</param>
    /// <param name="options">Optional gallery options.</param>
    /// <returns>Shared gallery manifest describing the generated artifacts and rich-content expectations.</returns>
    public static HtmlCapabilityGalleryManifest SaveHtmlCapabilityGallery(this ExcelDocument workbook, string directoryPath, ExcelHtmlCapabilityGalleryOptions? options = null) {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));
        if (string.IsNullOrWhiteSpace(directoryPath)) throw new ArgumentException("Artifact directory cannot be empty.", nameof(directoryPath));

        options ??= new ExcelHtmlCapabilityGalleryOptions();
        Directory.CreateDirectory(directoryPath);

        string filePrefix = ToFilePrefix(options.ScenarioId);
        string semanticHtml = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            Theme = options.Theme,
            Title = options.Title
        });
        string visualHtml = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelVisualReview,
            Theme = options.Theme,
            Title = options.Title + " Visual Review",
            VisualOptions = options.VisualOptions
        });

        string semanticPath = Path.Combine(directoryPath, filePrefix + ".semantic.html");
        string visualPath = Path.Combine(directoryPath, filePrefix + ".visual.html");
        string manifestPath = Path.Combine(directoryPath, filePrefix + ".manifest.md");
        string manifestJsonPath = Path.Combine(directoryPath, filePrefix + ".manifest.json");

        var scenario = new HtmlCapabilityGalleryScenario(
            options.ScenarioId,
            options.Title,
            "Excel HTML",
            "Validates Excel semantic table HTML and visual SVG review HTML for rich workbook content.");
        var result = new HtmlCapabilityGalleryResult(scenario);
        result.AddArtifact(HtmlCapabilityGalleryArtifact.WriteTextFile("semantic", "semantic-html", semanticPath, "text/html", semanticHtml));
        result.AddArtifact(HtmlCapabilityGalleryArtifact.WriteTextFile("visual", "visual-html", visualPath, "text/html", visualHtml));

        if (visualHtml.Contains("data-officeimo-visual-proof=\"comment-callout\"", StringComparison.Ordinal)) {
            result.Diagnostics.Add(
                "OfficeIMO.Excel.Html",
                "ExcelCommentVisualReviewRendered",
                "Excel comment bodies are visible in HTML visual review as dependency-free callout/list proof over the shared Drawing SVG export.",
                HtmlDiagnosticSeverity.Info);
        }

        if (semanticHtml.Contains("officeimo-chart-data", StringComparison.Ordinal)) {
            result.Diagnostics.Add(
                "OfficeIMO.Excel.Html",
                "ExcelChartSemanticDataPreserved",
                "Excel chart categories, series, and values were written as semantic HTML chart data.",
                HtmlDiagnosticSeverity.Info);
        }

        HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(semanticHtml, visualHtml);
        HtmlResourceManifest resources = HtmlResourcePipeline.BuildManifest(semanticHtml + visualHtml);
        var manifest = new HtmlCapabilityGalleryManifest(
            result,
            HtmlConversionProfile.PositionedReview,
            score,
            resources,
            CreateExpectations(workbook),
            new[] {
                OfficeHtmlConversionProfile.ExcelSemanticTables,
                OfficeHtmlConversionProfile.ExcelVisualReview
            });

        string manifestMarkdown = HtmlCapabilityGalleryManifestWriter.ToMarkdown(manifest);
        string manifestJson = HtmlCapabilityGalleryManifestJsonWriter.ToJson(manifest);
        HtmlCapabilityGalleryArtifact.WriteTextFile("manifest-md", "manifest-markdown", manifestPath, "text/markdown", manifestMarkdown);
        HtmlCapabilityGalleryArtifact.WriteTextFile("manifest-json", "manifest-json", manifestJsonPath, "application/json", manifestJson);

        return manifest;
    }

    private static IReadOnlyList<HtmlCapabilityGalleryExpectation> CreateExpectations(ExcelDocument workbook) {
        var expectations = new List<HtmlCapabilityGalleryExpectation>();
        bool hasUsedCells = false;
        bool hasFormulas = false;
        bool hasComments = false;
        bool hasCharts = false;
        bool hasImages = false;

        foreach (ExcelSheet sheet in workbook.Sheets) {
            hasUsedCells |= !string.Equals(sheet.GetUsedRangeA1(), "A1", StringComparison.OrdinalIgnoreCase) || sheet.TryGetCellText(1, 1, out string text) && text.Length > 0;
            hasFormulas |= sheet.GetFormulaCells().Count > 0;
            hasComments |= sheet.GetComments().Count > 0;
            hasCharts |= sheet.Charts.Any();
            hasImages |= sheet.Images.Any();
        }

        if (hasUsedCells) {
            expectations.Add(new HtmlCapabilityGalleryExpectation("worksheet tables", HtmlCapabilityGalleryExpectationOutcome.Preserved, "semantic HTML contains worksheet table cells"));
        }

        if (hasFormulas) {
            expectations.Add(new HtmlCapabilityGalleryExpectation("formulas", HtmlCapabilityGalleryExpectationOutcome.Preserved, "semantic HTML contains formula inventory"));
        }

        if (hasComments) {
            expectations.Add(new HtmlCapabilityGalleryExpectation("comments", HtmlCapabilityGalleryExpectationOutcome.VisualProof, "semantic inventory plus visual HTML comment callout/list proof"));
        }

        if (hasCharts) {
            expectations.Add(new HtmlCapabilityGalleryExpectation("charts", HtmlCapabilityGalleryExpectationOutcome.Preserved, "semantic chart data table plus visual SVG artifact"));
        }

        if (hasImages) {
            expectations.Add(new HtmlCapabilityGalleryExpectation("images", HtmlCapabilityGalleryExpectationOutcome.VisualProof, "semantic data URI preview plus visual SVG artifact"));
        }

        expectations.Add(new HtmlCapabilityGalleryExpectation("visual renderer owner", HtmlCapabilityGalleryExpectationOutcome.Reported, "visual HTML declares OfficeIMO.Drawing as the rendering owner"));
        return expectations;
    }

    private static string ToFilePrefix(string scenarioId) {
        var builder = new StringBuilder();
        foreach (char ch in scenarioId ?? string.Empty) {
            builder.Append(char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' ? ch : '-');
        }

        return builder.Length == 0 ? "excel-html-gallery" : builder.ToString();
    }
}
