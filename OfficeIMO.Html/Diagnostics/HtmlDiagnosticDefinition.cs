namespace OfficeIMO.Html;

/// <summary>
/// Support-facing metadata for a stable OfficeIMO HTML diagnostic code.
/// </summary>
public sealed class HtmlDiagnosticDefinition {
    internal HtmlDiagnosticDefinition(string code, string category, HtmlDiagnosticSeverity defaultSeverity, string explanation, string remediation) {
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Category = category ?? throw new ArgumentNullException(nameof(category));
        DefaultSeverity = defaultSeverity;
        Explanation = explanation ?? throw new ArgumentNullException(nameof(explanation));
        Remediation = remediation ?? throw new ArgumentNullException(nameof(remediation));
    }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Support category for grouping diagnostics.</summary>
    public string Category { get; }

    /// <summary>Default severity when an adapter does not provide a more specific severity.</summary>
    public HtmlDiagnosticSeverity DefaultSeverity { get; }

    /// <summary>Caller-facing explanation of what happened.</summary>
    public string Explanation { get; }

    /// <summary>Recommended remediation.</summary>
    public string Remediation { get; }
}
