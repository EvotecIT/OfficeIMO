using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Capability-gallery helpers for PowerPoint HTML exports.
/// </summary>
public static class PowerPointHtmlCapabilityGalleryExtensions {
    /// <summary>
    /// Saves semantic and visual PowerPoint HTML artifacts plus shared manifest files.
    /// </summary>
    /// <param name="presentation">Presentation to export.</param>
    /// <param name="directoryPath">Directory where artifacts will be written.</param>
    /// <param name="options">Optional gallery options.</param>
    /// <returns>Shared gallery manifest describing the generated artifacts and rich-content expectations.</returns>
    public static HtmlCapabilityGalleryManifest SaveHtmlCapabilityGallery(this PptCore.PowerPointPresentation presentation, string directoryPath, PowerPointHtmlCapabilityGalleryOptions? options = null) {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        if (string.IsNullOrWhiteSpace(directoryPath)) throw new ArgumentException("Artifact directory cannot be empty.", nameof(directoryPath));

        options ??= new PowerPointHtmlCapabilityGalleryOptions();
        Directory.CreateDirectory(directoryPath);

        string filePrefix = ToFilePrefix(options.ScenarioId);
        string semanticHtml = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides,
            Theme = options.Theme,
            Title = options.Title,
            IncludeHiddenSlides = options.IncludeHiddenSlides,
            IncludeNotes = options.IncludeNotes,
            IncludeTables = options.IncludeTables
        });
        string visualHtml = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointVisualReview,
            Theme = options.Theme,
            Title = options.Title + " Visual Review",
            IncludeHiddenSlides = options.IncludeHiddenSlides,
            IncludeNotes = options.IncludeNotes,
            IncludeTables = options.IncludeTables
        });

        string semanticPath = Path.Combine(directoryPath, filePrefix + ".semantic.html");
        string visualPath = Path.Combine(directoryPath, filePrefix + ".visual.html");
        string manifestPath = Path.Combine(directoryPath, filePrefix + ".manifest.md");
        string manifestJsonPath = Path.Combine(directoryPath, filePrefix + ".manifest.json");

        var scenario = new HtmlCapabilityGalleryScenario(
            options.ScenarioId,
            options.Title,
            "PowerPoint HTML",
            "Validates PowerPoint semantic slide HTML and positioned visual review HTML for rich presentation content.");
        var result = new HtmlCapabilityGalleryResult(scenario);
        result.AddArtifact(HtmlCapabilityGalleryArtifact.WriteTextFile("semantic", "semantic-html", semanticPath, "text/html", semanticHtml));
        result.AddArtifact(HtmlCapabilityGalleryArtifact.WriteTextFile("visual", "visual-html", visualPath, "text/html", visualHtml));

        if (visualHtml.Contains("officeimo-chart-rendered", StringComparison.Ordinal)) {
            result.Diagnostics.Add(
                "OfficeIMO.PowerPoint.Html",
                "PowerPointChartVisualReviewRendered",
                "PowerPoint chart visual review was rendered through the shared OfficeIMO.Drawing chart renderer.",
                HtmlDiagnosticSeverity.Info);
        }

        if (semanticHtml.Contains("officeimo-chart-data", StringComparison.Ordinal)) {
            result.Diagnostics.Add(
                "OfficeIMO.PowerPoint.Html",
                "PowerPointChartSemanticDataPreserved",
                "PowerPoint chart categories, series, and values were written as semantic HTML chart data.",
                HtmlDiagnosticSeverity.Info);
        }

        if (visualHtml.Contains("class=\"officeimo-shape-placeholder officeimo-chart-placeholder\"", StringComparison.Ordinal)) {
            result.Diagnostics.Add(
                "OfficeIMO.PowerPoint.Html",
                "PowerPointChartVisualPlaceholder",
                "PowerPoint chart visual rendering fell back to a positioned review placeholder; chart data is preserved as snapshot metadata.",
                HtmlDiagnosticSeverity.Warning);
        }

        HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(semanticHtml, visualHtml);
        HtmlResourceManifest resources = HtmlResourcePipeline.BuildManifest(semanticHtml + visualHtml);
        var manifest = new HtmlCapabilityGalleryManifest(
            result,
            HtmlConversionProfile.PositionedReview,
            score,
            resources,
            CreateExpectations(presentation, options),
            new[] {
                OfficeHtmlConversionProfile.PowerPointSemanticSlides,
                OfficeHtmlConversionProfile.PowerPointVisualReview
            });

        string manifestMarkdown = HtmlCapabilityGalleryManifestWriter.ToMarkdown(manifest);
        string manifestJson = HtmlCapabilityGalleryManifestJsonWriter.ToJson(manifest);
        HtmlCapabilityGalleryArtifact.WriteTextFile("manifest-md", "manifest-markdown", manifestPath, "text/markdown", manifestMarkdown);
        HtmlCapabilityGalleryArtifact.WriteTextFile("manifest-json", "manifest-json", manifestJsonPath, "application/json", manifestJson);

        return manifest;
    }

    private static IReadOnlyList<HtmlCapabilityGalleryExpectation> CreateExpectations(PptCore.PowerPointPresentation presentation, PowerPointHtmlCapabilityGalleryOptions options) {
        var expectations = new List<HtmlCapabilityGalleryExpectation>();
        IEnumerable<PptCore.PowerPointSlide> slides = options.IncludeHiddenSlides
            ? presentation.Slides
            : presentation.Slides.Where(slide => !slide.Hidden);
        List<PptCore.PowerPointSlide> slideList = slides.ToList();

        if (slideList.Any(slide => slide.TextBoxes.Any(textBox => !string.IsNullOrWhiteSpace(textBox.Text)))) {
            expectations.Add(new HtmlCapabilityGalleryExpectation("text boxes", HtmlCapabilityGalleryExpectationOutcome.Preserved, "semantic HTML contains extracted slide text"));
        }

        if (options.IncludeTables && slideList.Any(slide => slide.Tables.Any())) {
            expectations.Add(new HtmlCapabilityGalleryExpectation("tables", HtmlCapabilityGalleryExpectationOutcome.Preserved, "semantic and visual HTML contain table cells"));
        }

        if (slideList.Any(slide => slide.Pictures.Any())) {
            expectations.Add(new HtmlCapabilityGalleryExpectation("pictures", HtmlCapabilityGalleryExpectationOutcome.VisualProof, "semantic inventory plus positioned image data URI"));
        }

        if (slideList.Any(slide => slide.Charts.Any())) {
            expectations.Add(new HtmlCapabilityGalleryExpectation("charts", HtmlCapabilityGalleryExpectationOutcome.Preserved, "semantic chart data table plus shared Drawing SVG visual proof when supported"));
        }

        expectations.Add(new HtmlCapabilityGalleryExpectation("positioned review", HtmlCapabilityGalleryExpectationOutcome.VisualProof, "visual HTML declares positioned-review boundary"));
        return expectations;
    }

    private static string ToFilePrefix(string scenarioId) {
        var builder = new StringBuilder();
        foreach (char ch in scenarioId ?? string.Empty) {
            builder.Append(char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' ? ch : '-');
        }

        return builder.Length == 0 ? "powerpoint-html-gallery" : builder.ToString();
    }
}
