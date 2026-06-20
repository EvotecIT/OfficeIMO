namespace OfficeIMO.Html;

/// <summary>
/// Converts resource manifests into grouped dependency plans for adapters and reports.
/// </summary>
public static class HtmlResourceDependencyPlanner {
    /// <summary>
    /// Creates a dependency plan from a resource manifest.
    /// </summary>
    public static HtmlResourceDependencyPlan Create(HtmlResourceManifest manifest) {
        if (manifest == null) {
            throw new ArgumentNullException(nameof(manifest));
        }

        var allowed = manifest.Resources.Where(resource => resource.IsAllowed).ToList().AsReadOnly();
        var blocked = manifest.Resources.Where(resource => !resource.IsAllowed).ToList().AsReadOnly();
        var summaries = manifest.Resources
            .GroupBy(resource => resource.Kind)
            .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
            .Select(group => new HtmlResourceDependencySummary(
                group.Key,
                group.Count(),
                group.Count(resource => resource.IsAllowed),
                group.Count(resource => !resource.IsAllowed)))
            .ToList()
            .AsReadOnly();

        return new HtmlResourceDependencyPlan(allowed, blocked, summaries);
    }
}
