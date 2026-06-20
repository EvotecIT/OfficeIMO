namespace OfficeIMO.Html;

/// <summary>
/// Manifest of resource dependencies discovered in HTML input.
/// </summary>
public sealed class HtmlResourceManifest {
    private readonly List<HtmlResourceReference> _resources = new List<HtmlResourceReference>();
    private readonly HtmlDiagnosticReport _diagnostics = new HtmlDiagnosticReport();

    /// <summary>Resources in document order.</summary>
    public IReadOnlyList<HtmlResourceReference> Resources => _resources;

    /// <summary>Diagnostics produced while planning resources.</summary>
    public HtmlDiagnosticReport Diagnostics => _diagnostics;

    /// <summary>Number of allowed resources.</summary>
    public int AllowedCount => _resources.Count(resource => resource.IsAllowed);

    /// <summary>Number of blocked resources.</summary>
    public int BlockedCount => _resources.Count(resource => !resource.IsAllowed);

    internal void Add(HtmlResourceReference resource) {
        _resources.Add(resource);
        if (!resource.IsAllowed) {
            _diagnostics.Add(
                "OfficeIMO.Html",
                resource.DiagnosticCode.Length == 0 ? "HtmlResourceRejectedByPolicy" : resource.DiagnosticCode,
                "HTML resource was rejected by the configured URL policy.",
                HtmlDiagnosticSeverity.Warning,
                resource.Source,
                resource.ElementName + "[" + resource.AttributeName + "]");
        }
    }
}
