namespace OfficeIMO.Html;

/// <summary>
/// Computed-style snapshot for one HTML element.
/// </summary>
public sealed class HtmlComputedStyle {
    private static readonly IReadOnlyDictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace> EmptyCascadeTraces =
        new System.Collections.ObjectModel.ReadOnlyDictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace>(
            new Dictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace>(HtmlCssPropertyNameComparer.Instance));
    private readonly Dictionary<string, string> _properties;
    private readonly IReadOnlyDictionary<string, string> _readOnlyProperties;
    private readonly HashSet<string> _inheritedProperties;
    private readonly HashSet<string> _resetProperties;
    private readonly HashSet<string> _originRevertedProperties;
    private readonly HashSet<string> _specifiedProperties;
    private readonly Dictionary<string, HtmlCssCascadePriority> _cascadePriorities;
    private readonly IReadOnlyDictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace> _cascadeTraces;

    internal HtmlComputedStyle(
        IDictionary<string, string> properties,
        IEnumerable<string>? inheritedProperties = null,
        IEnumerable<string>? resetProperties = null,
        IEnumerable<string>? originRevertedProperties = null,
        IEnumerable<string>? specifiedProperties = null,
        IDictionary<string, HtmlCssCascadePriority>? cascadePriorities = null,
        IDictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace>? cascadeTraces = null) {
        _properties = new Dictionary<string, string>(properties ?? throw new ArgumentNullException(nameof(properties)), HtmlCssPropertyNameComparer.Instance);
        _readOnlyProperties = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(_properties);
        _inheritedProperties = new HashSet<string>(inheritedProperties ?? Array.Empty<string>(), HtmlCssPropertyNameComparer.Instance);
        _resetProperties = new HashSet<string>(resetProperties ?? Array.Empty<string>(), HtmlCssPropertyNameComparer.Instance);
        _originRevertedProperties = new HashSet<string>(originRevertedProperties ?? Array.Empty<string>(), HtmlCssPropertyNameComparer.Instance);
        _specifiedProperties = new HashSet<string>(specifiedProperties ?? Array.Empty<string>(), HtmlCssPropertyNameComparer.Instance);
        _cascadePriorities = new Dictionary<string, HtmlCssCascadePriority>(cascadePriorities ?? new Dictionary<string, HtmlCssCascadePriority>(), HtmlCssPropertyNameComparer.Instance);
        _cascadeTraces = cascadeTraces == null || cascadeTraces.Count == 0
            ? EmptyCascadeTraces
            : new System.Collections.ObjectModel.ReadOnlyDictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace>(
                new Dictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace>(cascadeTraces, HtmlCssPropertyNameComparer.Instance));
    }

    private HtmlComputedStyle(
        Dictionary<string, string> properties,
        HashSet<string> inheritedProperties,
        HashSet<string> resetProperties,
        HashSet<string> originRevertedProperties,
        HashSet<string> specifiedProperties,
        Dictionary<string, HtmlCssCascadePriority> cascadePriorities,
        Dictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace>? cascadeTraces) {
        _properties = properties;
        _readOnlyProperties = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(_properties);
        _inheritedProperties = inheritedProperties;
        _resetProperties = resetProperties;
        _originRevertedProperties = originRevertedProperties;
        _specifiedProperties = specifiedProperties;
        _cascadePriorities = cascadePriorities;
        _cascadeTraces = cascadeTraces == null || cascadeTraces.Count == 0 ? EmptyCascadeTraces : cascadeTraces;
    }

    internal static HtmlComputedStyle FromOwnedCollections(
        Dictionary<string, string> properties,
        HashSet<string> inheritedProperties,
        HashSet<string> resetProperties,
        HashSet<string> originRevertedProperties,
        HashSet<string> specifiedProperties,
        Dictionary<string, HtmlCssCascadePriority> cascadePriorities,
        Dictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace>? cascadeTraces) =>
        new HtmlComputedStyle(properties, inheritedProperties, resetProperties, originRevertedProperties,
            specifiedProperties, cascadePriorities, cascadeTraces);

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

    /// <summary>
    /// Returns the provider-independent typed value for a computed property when its current
    /// OfficeIMO grammar slice accepts the value. Layout-dependent percentages remain in the
    /// returned expression until a caller supplies a used-value resolution context.
    /// </summary>
    public bool TryGetTypedValue(string propertyName, out OfficeIMO.Html.Css.HtmlCssPropertyValue? value) {
        value = null;
        string computed = GetValue(propertyName);
        if (computed.Length == 0) return false;
        OfficeIMO.Html.Css.HtmlCssPropertyParseResult parsed =
            OfficeIMO.Html.Css.HtmlCssPropertyParser.Parse(propertyName, computed);
        if (parsed.Status != OfficeIMO.Html.Css.HtmlCssPropertyParseStatus.Parsed) return false;
        value = parsed.Value;
        return value != null;
    }

    internal bool IsInheritedValue(string propertyName) =>
        !string.IsNullOrWhiteSpace(propertyName) && _inheritedProperties.Contains(propertyName.Trim());

    internal bool IsImplicitlyInheritedValue(string propertyName) {
        if (string.IsNullOrWhiteSpace(propertyName)) return false;
        string name = propertyName.Trim();
        return _inheritedProperties.Contains(name)
            && _cascadePriorities.TryGetValue(name, out HtmlCssCascadePriority priority)
            && priority.IsImplicitInheritance;
    }

    /// <summary>Returns an owned cascade explanation for a property in the implemented grammar slice.</summary>
    public bool TryGetCascadeTrace(string propertyName, out OfficeIMO.Html.Css.HtmlCssCascadeTrace? trace) {
        if (string.IsNullOrWhiteSpace(propertyName)) {
            trace = null;
            return false;
        }
        return _cascadeTraces.TryGetValue(propertyName.Trim(), out trace);
    }

    /// <summary>Gets an owned cascade explanation, or null when the property's trace slice is not implemented.</summary>
    public OfficeIMO.Html.Css.HtmlCssCascadeTrace? GetCascadeTrace(string propertyName) =>
        TryGetCascadeTrace(propertyName, out OfficeIMO.Html.Css.HtmlCssCascadeTrace? trace) ? trace : null;

    internal bool IsResetValue(string propertyName) =>
        !string.IsNullOrWhiteSpace(propertyName) && _resetProperties.Contains(propertyName.Trim());

    internal bool IsOriginRevertedValue(string propertyName) =>
        !string.IsNullOrWhiteSpace(propertyName) && _originRevertedProperties.Contains(propertyName.Trim());

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
            new HashSet<string>(_originRevertedProperties, HtmlCssPropertyNameComparer.Instance),
            new HashSet<string>(_specifiedProperties, HtmlCssPropertyNameComparer.Instance),
            cascadePriorities,
            _cascadeTraces.Count == 0
                ? null
                : _cascadeTraces.ToDictionary(
                    pair => pair.Key,
                    pair => pair.Value,
                    HtmlCssPropertyNameComparer.Instance));
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

    internal bool Inherited { get; }
    internal bool Important { get; }
    internal bool Inline { get; }
    internal CascadeLayerOrder? LayerOrder { get; }
    internal int Ids { get; }
    internal int Classes { get; }
    internal int Elements { get; }
    internal int RuleOrder { get; }
    internal int DeclarationOrder { get; }
    internal bool IsImplicitInheritance =>
        Inherited && !Important && !Inline && LayerOrder == null
        && Ids < 0 && Classes < 0 && Elements < 0 && RuleOrder < 0 && DeclarationOrder < 0;

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
