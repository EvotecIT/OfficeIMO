namespace OfficeIMO.Html;

/// <summary>
/// Describes one HTML capability-gallery scenario and its user-facing purpose.
/// </summary>
public sealed class HtmlCapabilityGalleryScenario {
    /// <summary>
    /// Creates a capability-gallery scenario.
    /// </summary>
    /// <param name="id">Stable scenario identifier used in artifact names and manifests.</param>
    /// <param name="title">Human-readable scenario title.</param>
    /// <param name="category">Scenario category, such as <c>Word HTML</c> or <c>HTML PDF</c>.</param>
    /// <param name="description">Short explanation of the capability being proved.</param>
    public HtmlCapabilityGalleryScenario(string id, string title, string category, string description) {
        Id = id ?? throw new ArgumentNullException(nameof(id));
        Title = title ?? throw new ArgumentNullException(nameof(title));
        Category = category ?? throw new ArgumentNullException(nameof(category));
        Description = description ?? throw new ArgumentNullException(nameof(description));
    }

    /// <summary>
    /// Stable scenario identifier used in artifact names and manifests.
    /// </summary>
    public string Id { get; }

    /// <summary>
    /// Human-readable scenario title.
    /// </summary>
    public string Title { get; }

    /// <summary>
    /// Scenario category, such as <c>Word HTML</c> or <c>HTML PDF</c>.
    /// </summary>
    public string Category { get; }

    /// <summary>
    /// Short explanation of the capability being proved.
    /// </summary>
    public string Description { get; }
}
