namespace OfficeIMO.Html;

/// <summary>
/// Shared catalog of OfficeIMO HTML conversion profile contracts.
/// </summary>
public static class HtmlConversionProfileContracts {
    private static readonly IReadOnlyList<HtmlConversionProfileContract> Contracts = new List<HtmlConversionProfileContract> {
        new HtmlConversionProfileContract(
            HtmlConversionProfile.Semantic,
            "Semantic",
            "Editable office documents, clean HTML export, accessible reports, and deterministic round-trips.",
            "Preserve document structure and meaning first; simplify browser-only layout where needed.",
            new[] { "headings", "paragraphs", "lists", "tables", "links", "images", "forms-as-document-fields", "language-and-direction-metadata" },
            new[] { "inline-style", "document-style-rules", "text-formatting", "table-borders", "spacing", "colors", "direction" },
            new[] { "URL policy enforcement", "responsive image candidate selection", "resource manifest reporting" },
            new[] { "unsupported CSS warnings", "resource policy diagnostics", "accessibility advisories", "round-trip score evidence" }),
        new HtmlConversionProfileContract(
            HtmlConversionProfile.Document,
            "Document",
            "Business documents, invoices, contracts, generated reports, and HTML intended to become DOCX/PDF artifacts.",
            "Balance editability and visual fidelity for common document layouts.",
            new[] { "semantic sections", "headers-and-footers", "tables-with-spans", "captions", "figures", "form-controls", "comments" },
            new[] { "cascade snapshot", "selector matching", "font-and-color inheritance", "table layout hints", "print-friendly spacing" },
            new[] { "bounded downloads", "content-type validation", "byte budgets", "base URI resolution", "blocked resource reporting" },
            new[] { "diagnostic catalog lookup", "shared report aggregation", "gallery manifest diagnostics" }),
        new HtmlConversionProfileContract(
            HtmlConversionProfile.HighFidelityPrint,
            "High Fidelity Print",
            "Print/PDF lanes, visual review, and workflows where page appearance matters more than editable structure.",
            "Expose layout preservation as an explicit ambition while reporting fallbacks and unsupported browser features.",
            new[] { "print sections", "positioning hints", "backgrounds", "page-breaks", "complex tables", "media-heavy content" },
            new[] { "computed-style capture", "media intent metadata", "layout-affecting declarations", "resource dependency graph" },
            new[] { "complete resource inventory", "policy outcome per resource", "external dependency diagnostics" },
            new[] { "fidelity score", "layout fallback diagnostics", "unsupported high-fidelity feature diagnostics" }),
        new HtmlConversionProfileContract(
            HtmlConversionProfile.PositionedReview,
            "Positioned Review",
            "PDF readback, page previews, and diagnostic review lanes where source geometry needs to remain inspectable in HTML.",
            "Preserve review geometry and source anchors while clearly avoiding editable document reconstruction claims.",
            new[] { "page wrappers", "positioned text blocks", "positioned images", "link frames", "form field frames", "source anchors" },
            new[] { "absolute positioning", "page dimensions", "safe overlay styles", "review-only visual hints" },
            new[] { "resource inventory", "safe link handling", "image placeholder or embedding policy", "source coordinate reporting" },
            new[] { "geometry simplification diagnostics", "unsafe link diagnostics", "missing resource diagnostics", "no-editable-reconstruction boundary" })
    }.AsReadOnly();

    /// <summary>
    /// Gets all shared profile contracts.
    /// </summary>
    public static IReadOnlyList<HtmlConversionProfileContract> All => Contracts;

    /// <summary>
    /// Gets a shared profile contract by profile identifier.
    /// </summary>
    public static HtmlConversionProfileContract Get(HtmlConversionProfile profile) {
        foreach (HtmlConversionProfileContract contract in Contracts) {
            if (contract.Profile == profile) {
                return contract;
            }
        }

        throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unknown HTML conversion profile.");
    }
}
