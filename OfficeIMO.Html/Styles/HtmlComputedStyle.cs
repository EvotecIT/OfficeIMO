namespace OfficeIMO.Html;

/// <summary>
/// Computed-style snapshot for one HTML element.
/// </summary>
public sealed class HtmlComputedStyle {
    private readonly Dictionary<string, string> _properties;
    private readonly IReadOnlyDictionary<string, string> _readOnlyProperties;

    internal HtmlComputedStyle(IDictionary<string, string> properties) {
        _properties = new Dictionary<string, string>(properties ?? throw new ArgumentNullException(nameof(properties)), StringComparer.OrdinalIgnoreCase);
        _readOnlyProperties = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(_properties);
    }

    /// <summary>All computed properties known to the lightweight style engine.</summary>
    public IReadOnlyDictionary<string, string> Properties => _readOnlyProperties;

    /// <summary>Gets a computed property value or an empty string when no value is known.</summary>
    public string GetValue(string propertyName) {
        if (string.IsNullOrWhiteSpace(propertyName)) {
            return string.Empty;
        }

        return _properties.TryGetValue(propertyName.Trim(), out string? value) ? value : string.Empty;
    }
}
