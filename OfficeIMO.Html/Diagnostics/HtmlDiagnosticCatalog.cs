namespace OfficeIMO.Html;

/// <summary>
/// Catalog of stable OfficeIMO HTML diagnostics and support remediation text.
/// </summary>
public static class HtmlDiagnosticCatalog {
    private static readonly IReadOnlyDictionary<string, HtmlDiagnosticDefinition> Definitions = new Dictionary<string, HtmlDiagnosticDefinition>(StringComparer.OrdinalIgnoreCase) {
        ["HtmlCommentSkipped"] = new HtmlDiagnosticDefinition(
            "HtmlCommentSkipped",
            "ContentSimplification",
            HtmlDiagnosticSeverity.Info,
            "An HTML comment was omitted from generated document content.",
            "Enable HTML comment import when comments are part of the expected document contract, or keep comments as source-only metadata."),
        ["ImageResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "ImageResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "An image candidate was rejected before loading because its URI is not allowed by policy.",
            "Allow the URI scheme or host for trusted inputs, embed the image as data URI, or provide a local resource resolver."),
        ["StylesheetResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "StylesheetResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A stylesheet was rejected before loading because its URI is not allowed by policy.",
            "Use caller-provided stylesheet contents for untrusted HTML, or allow the stylesheet scheme and host for trusted documents."),
        ["HyperlinkRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "HyperlinkRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A hyperlink target was rejected because its URI is not allowed by policy.",
            "Use http, https, mailto, or a caller-approved scheme instead of script or local file targets."),
        ["ScriptResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "ScriptResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A script dependency was rejected before loading because its URI is not allowed by policy.",
            "Use caller-provided script handling for trusted automation scenarios, or remove script dependencies from document-oriented HTML inputs."),
        ["MediaResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "MediaResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A media dependency was rejected before loading because its URI is not allowed by policy.",
            "Allow trusted media hosts explicitly, package approved media with the input, or provide a local resource resolver."),
        ["FontResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "FontResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A font dependency was rejected before loading because its URI is not allowed by policy.",
            "Use packaged fonts from trusted locations or allow approved font hosts in the URL policy."),
        ["UnsupportedCssDeclaration"] = new HtmlDiagnosticDefinition(
            "UnsupportedCssDeclaration",
            "CssFidelity",
            HtmlDiagnosticSeverity.Warning,
            "A CSS declaration could not be mapped to the target document model.",
            "Prefer document-friendly CSS or route visual-first workloads through the high-fidelity print profile."),
        ["HtmlResourceRejectedByPolicy"] = new HtmlDiagnosticDefinition(
            "HtmlResourceRejectedByPolicy",
            "ResourcePolicy",
            HtmlDiagnosticSeverity.Warning,
            "A resource dependency was rejected before loading because its URI is not allowed by policy.",
            "Adjust the URL policy only for trusted sources, or package the dependency with the HTML input.")
    };

    /// <summary>
    /// Gets all known diagnostic definitions.
    /// </summary>
    public static IReadOnlyDictionary<string, HtmlDiagnosticDefinition> All => Definitions;

    /// <summary>
    /// Looks up support metadata for a diagnostic code.
    /// </summary>
    public static bool TryGet(string code, out HtmlDiagnosticDefinition definition) {
        if (string.IsNullOrWhiteSpace(code)) {
            definition = null!;
            return false;
        }

        HtmlDiagnosticDefinition? found;
        if (Definitions.TryGetValue(code.Trim(), out found)) {
            definition = found;
            return true;
        }

        definition = null!;
        return false;
    }

    /// <summary>
    /// Gets support metadata for a diagnostic code, or a generic definition when the code is unknown.
    /// </summary>
    public static HtmlDiagnosticDefinition GetOrCreateGeneric(string code) {
        if (TryGet(code, out HtmlDiagnosticDefinition definition)) {
            return definition;
        }

        return new HtmlDiagnosticDefinition(
            string.IsNullOrWhiteSpace(code) ? "HtmlDiagnostic" : code.Trim(),
            "General",
            HtmlDiagnosticSeverity.Warning,
            "The HTML workflow emitted a diagnostic that is not yet cataloged.",
            "Use the diagnostic source and detail fields to decide whether input, policy, or converter support should be adjusted.");
    }
}
