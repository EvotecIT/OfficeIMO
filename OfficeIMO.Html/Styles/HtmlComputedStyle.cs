namespace OfficeIMO.Html;

/// <summary>
/// Computed-style snapshot for one HTML element.
/// </summary>
public sealed class HtmlComputedStyle {
    private readonly Dictionary<string, string> _properties;
    private readonly IReadOnlyDictionary<string, string> _readOnlyProperties;
    private readonly HashSet<string> _inheritedProperties;
    private readonly HashSet<string> _resetProperties;
    private readonly HashSet<string> _specifiedProperties;

    internal HtmlComputedStyle(
        IDictionary<string, string> properties,
        IEnumerable<string>? inheritedProperties = null,
        IEnumerable<string>? resetProperties = null,
        IEnumerable<string>? specifiedProperties = null) {
        _properties = new Dictionary<string, string>(properties ?? throw new ArgumentNullException(nameof(properties)), HtmlCssPropertyNameComparer.Instance);
        _readOnlyProperties = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(_properties);
        _inheritedProperties = new HashSet<string>(inheritedProperties ?? Array.Empty<string>(), HtmlCssPropertyNameComparer.Instance);
        _resetProperties = new HashSet<string>(resetProperties ?? Array.Empty<string>(), HtmlCssPropertyNameComparer.Instance);
        _specifiedProperties = new HashSet<string>(specifiedProperties ?? Array.Empty<string>(), HtmlCssPropertyNameComparer.Instance);
    }

    private HtmlComputedStyle(
        Dictionary<string, string> properties,
        HashSet<string> inheritedProperties,
        HashSet<string> resetProperties,
        HashSet<string> specifiedProperties) {
        _properties = properties;
        _readOnlyProperties = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(_properties);
        _inheritedProperties = inheritedProperties;
        _resetProperties = resetProperties;
        _specifiedProperties = specifiedProperties;
    }

    internal static HtmlComputedStyle FromOwnedCollections(
        Dictionary<string, string> properties,
        HashSet<string> inheritedProperties,
        HashSet<string> resetProperties,
        HashSet<string> specifiedProperties) =>
        new HtmlComputedStyle(properties, inheritedProperties, resetProperties, specifiedProperties);

    /// <summary>All computed properties known to the lightweight style engine.</summary>
    public IReadOnlyDictionary<string, string> Properties => _readOnlyProperties;

    /// <summary>
    /// Gets the resolved effective font size in points after relative units, percentages,
    /// and CSS font-size keywords have been evaluated against their inherited context.
    /// </summary>
    public double? ResolvedFontSizePoints { get; internal set; }

    /// <summary>Gets a computed property value or an empty string when no value is known.</summary>
    public string GetValue(string propertyName) {
        if (string.IsNullOrWhiteSpace(propertyName)) {
            return string.Empty;
        }

        return _properties.TryGetValue(propertyName.Trim(), out string? value) ? value : string.Empty;
    }

    internal bool IsInheritedValue(string propertyName) =>
        !string.IsNullOrWhiteSpace(propertyName) && _inheritedProperties.Contains(propertyName.Trim());

    internal bool IsResetValue(string propertyName) =>
        !string.IsNullOrWhiteSpace(propertyName) && _resetProperties.Contains(propertyName.Trim());

    internal bool IsSpecifiedValue(string propertyName) =>
        !string.IsNullOrWhiteSpace(propertyName) && _specifiedProperties.Contains(propertyName.Trim());
}

internal sealed class HtmlCssPropertyNameComparer : IEqualityComparer<string> {
    internal static HtmlCssPropertyNameComparer Instance { get; } = new HtmlCssPropertyNameComparer();

    private HtmlCssPropertyNameComparer() {
    }

    public bool Equals(string? x, string? y) {
        bool xIsCustom = x?.StartsWith("--", StringComparison.Ordinal) == true;
        bool yIsCustom = y?.StartsWith("--", StringComparison.Ordinal) == true;
        return xIsCustom || yIsCustom
            ? StringComparer.Ordinal.Equals(x, y)
            : StringComparer.OrdinalIgnoreCase.Equals(x, y);
    }

    public int GetHashCode(string value) =>
        value.StartsWith("--", StringComparison.Ordinal)
            ? StringComparer.Ordinal.GetHashCode(value)
            : StringComparer.OrdinalIgnoreCase.GetHashCode(value);
}
