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
    private readonly Dictionary<string, HtmlCssCascadePriority> _cascadePriorities;

    internal HtmlComputedStyle(
        IDictionary<string, string> properties,
        IEnumerable<string>? inheritedProperties = null,
        IEnumerable<string>? resetProperties = null,
        IEnumerable<string>? specifiedProperties = null,
        IDictionary<string, HtmlCssCascadePriority>? cascadePriorities = null) {
        _properties = new Dictionary<string, string>(properties ?? throw new ArgumentNullException(nameof(properties)), HtmlCssPropertyNameComparer.Instance);
        _readOnlyProperties = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(_properties);
        _inheritedProperties = new HashSet<string>(inheritedProperties ?? Array.Empty<string>(), HtmlCssPropertyNameComparer.Instance);
        _resetProperties = new HashSet<string>(resetProperties ?? Array.Empty<string>(), HtmlCssPropertyNameComparer.Instance);
        _specifiedProperties = new HashSet<string>(specifiedProperties ?? Array.Empty<string>(), HtmlCssPropertyNameComparer.Instance);
        _cascadePriorities = new Dictionary<string, HtmlCssCascadePriority>(cascadePriorities ?? new Dictionary<string, HtmlCssCascadePriority>(), HtmlCssPropertyNameComparer.Instance);
    }

    private HtmlComputedStyle(
        Dictionary<string, string> properties,
        HashSet<string> inheritedProperties,
        HashSet<string> resetProperties,
        HashSet<string> specifiedProperties,
        Dictionary<string, HtmlCssCascadePriority> cascadePriorities) {
        _properties = properties;
        _readOnlyProperties = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(_properties);
        _inheritedProperties = inheritedProperties;
        _resetProperties = resetProperties;
        _specifiedProperties = specifiedProperties;
        _cascadePriorities = cascadePriorities;
    }

    internal static HtmlComputedStyle FromOwnedCollections(
        Dictionary<string, string> properties,
        HashSet<string> inheritedProperties,
        HashSet<string> resetProperties,
        HashSet<string> specifiedProperties,
        Dictionary<string, HtmlCssCascadePriority> cascadePriorities) =>
        new HtmlComputedStyle(properties, inheritedProperties, resetProperties, specifiedProperties, cascadePriorities);

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

    internal bool ShouldOverride(string candidateProperty, string existingProperty) {
        if (!_cascadePriorities.TryGetValue(candidateProperty, out HtmlCssCascadePriority candidate)) return true;
        return !_cascadePriorities.TryGetValue(existingProperty, out HtmlCssCascadePriority existing)
            || candidate.OutranksOrEquals(existing);
    }

    internal HtmlComputedStyle WithMappedProperties(
        Dictionary<string, string> properties,
        Dictionary<string, HtmlCssCascadePriority> cascadePriorities) {
        var style = new HtmlComputedStyle(
            properties,
            new HashSet<string>(_inheritedProperties, HtmlCssPropertyNameComparer.Instance),
            new HashSet<string>(_resetProperties, HtmlCssPropertyNameComparer.Instance),
            new HashSet<string>(_specifiedProperties, HtmlCssPropertyNameComparer.Instance),
            cascadePriorities);
        style.ResolvedFontSizePoints = ResolvedFontSizePoints;
        return style;
    }

    internal Dictionary<string, HtmlCssCascadePriority> CopyCascadePriorities() =>
        new Dictionary<string, HtmlCssCascadePriority>(_cascadePriorities, HtmlCssPropertyNameComparer.Instance);

    internal bool TryGetCascadePriority(string propertyName, out HtmlCssCascadePriority priority) =>
        _cascadePriorities.TryGetValue(propertyName, out priority);
}

internal readonly struct HtmlCssCascadePriority {
    internal HtmlCssCascadePriority(
        bool inherited,
        bool important,
        bool inline,
        CascadeLayerOrder? layerOrder,
        int ids,
        int classes,
        int elements,
        int ruleOrder,
        int declarationOrder) {
        Inherited = inherited;
        Important = important;
        Inline = inline;
        LayerOrder = layerOrder;
        Ids = ids;
        Classes = classes;
        Elements = elements;
        RuleOrder = ruleOrder;
        DeclarationOrder = declarationOrder;
    }

    private bool Inherited { get; }
    private bool Important { get; }
    private bool Inline { get; }
    private CascadeLayerOrder? LayerOrder { get; }
    private int Ids { get; }
    private int Classes { get; }
    private int Elements { get; }
    private int RuleOrder { get; }
    private int DeclarationOrder { get; }

    internal bool OutranksOrEquals(HtmlCssCascadePriority existing) {
        if (existing.Inherited != Inherited) return !Inherited;
        if (Important != existing.Important) return Important;
        if (Important && Inline != existing.Inline) return Inline;
        if ((LayerOrder != null) != (existing.LayerOrder != null)) {
            return Important ? LayerOrder != null : LayerOrder == null;
        }
        if (LayerOrder != null && existing.LayerOrder != null) {
            int layerComparison = LayerOrder.CompareTo(existing.LayerOrder);
            if (layerComparison != 0) return Important ? layerComparison < 0 : layerComparison > 0;
        }
        if (Ids != existing.Ids) return Ids > existing.Ids;
        if (Classes != existing.Classes) return Classes > existing.Classes;
        if (Elements != existing.Elements) return Elements > existing.Elements;
        if (RuleOrder != existing.RuleOrder) return RuleOrder > existing.RuleOrder;
        return DeclarationOrder >= existing.DeclarationOrder;
    }
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
