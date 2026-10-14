namespace OfficeIMO.Html;

/// <summary>
/// Resource dependency discovered during OfficeIMO HTML resource planning.
/// </summary>
public sealed class HtmlResourceReference {
    internal HtmlResourceReference(HtmlResourceKind kind, string elementName, string attributeName, string source, string resolvedSource, bool isAllowed, string diagnosticCode, Uri? resolutionBaseUri, Func<string, string?>? resolvedUrlTransform) {
        Kind = kind;
        ElementName = elementName ?? string.Empty;
        AttributeName = attributeName ?? string.Empty;
        Source = source ?? string.Empty;
        ResolvedSource = resolvedSource ?? string.Empty;
        IsAllowed = isAllowed;
        DiagnosticCode = diagnosticCode ?? string.Empty;
        ResolutionBaseUri = resolutionBaseUri;
        PlanningUrlTransform = resolvedUrlTransform;
    }

    /// <summary>Resource kind.</summary>
    public HtmlResourceKind Kind { get; }

    /// <summary>Element name where the dependency was found.</summary>
    public string ElementName { get; }

    /// <summary>Attribute name where the dependency was found.</summary>
    public string AttributeName { get; }

    /// <summary>Raw source value.</summary>
    public string Source { get; }

    /// <summary>Resolved source value, empty when blocked or unresolved.</summary>
    public string ResolvedSource { get; }

    /// <summary>Whether the source passed URL policy evaluation.</summary>
    public bool IsAllowed { get; }

    /// <summary>Diagnostic code associated with blocked or degraded resource handling.</summary>
    public string DiagnosticCode { get; }

    // Retain the planner's original resolution context so operation policies can
    // authorize raw input without applying a rewriting transform twice.
    internal Uri? ResolutionBaseUri { get; }
    internal Func<string, string?>? PlanningUrlTransform { get; }
}
