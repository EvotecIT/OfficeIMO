namespace OfficeIMO.Html;

/// <summary>
/// Curated scenario catalog for market-facing OfficeIMO HTML examples.
/// </summary>
public static class HtmlMarketScenarioCatalog {
    private static readonly IReadOnlyList<HtmlMarketScenario> Scenarios = new List<HtmlMarketScenario> {
        new HtmlMarketScenario("invoice", "Invoice", HtmlConversionProfile.Document, new[] { "tables", "currency", "branding", "round-trip" }, "Convert billing HTML into editable DOCX and PDF-ready output."),
        new HtmlMarketScenario("quarterly-report", "Quarterly Report", HtmlConversionProfile.Document, new[] { "headings", "tables", "forms", "diagnostics" }, "Turn generated operational reports into validated office documents."),
        new HtmlMarketScenario("legal-contract", "Legal Contract", HtmlConversionProfile.Semantic, new[] { "headings", "numbered-lists", "comments", "links" }, "Preserve contract structure for editing, review, and audit trails."),
        new HtmlMarketScenario("email-render", "Email Render", HtmlConversionProfile.Semantic, new[] { "inline-styles", "images", "links", "resource-policy" }, "Ingest email-like HTML safely while reporting blocked resources."),
        new HtmlMarketScenario("dashboard-print", "Dashboard Print", HtmlConversionProfile.HighFidelityPrint, new[] { "cards", "charts", "computed-style", "print-layout" }, "Create a visual-first review artifact from dashboard HTML."),
        new HtmlMarketScenario("multilingual-bidi", "Multilingual BiDi", HtmlConversionProfile.Document, new[] { "language", "direction", "tables", "fonts" }, "Validate right-to-left and mixed-language document conversion.")
    }.AsReadOnly();

    /// <summary>Gets all curated market scenarios.</summary>
    public static IReadOnlyList<HtmlMarketScenario> All => Scenarios;

    /// <summary>Gets a scenario by id.</summary>
    public static HtmlMarketScenario Get(string id) {
        foreach (HtmlMarketScenario scenario in Scenarios) {
            if (string.Equals(scenario.Id, id, StringComparison.OrdinalIgnoreCase)) {
                return scenario;
            }
        }

        throw new ArgumentException("Unknown HTML market scenario '" + id + "'.", nameof(id));
    }
}
