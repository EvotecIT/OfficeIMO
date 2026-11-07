namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private const int MaximumDeclarationFacts = 1024;
    private readonly Dictionary<object, DeclarationFacts> _declarationFacts = new();

    /// <summary>Reuses only immutable declaration predicates; character metrics and used values still resolve in the current context.</summary>
    private DeclarationFacts GetDeclarationFacts(HtmlComputedStyle computed) {
        object identity = computed.DeclarationIdentity;
        if (_declarationFacts.TryGetValue(identity, out DeclarationFacts facts)) return facts;
        // Computed ordinary declarations already contain resolved var() values;
        // unused inherited custom payloads cannot affect character-unit sizing.
        facts = new DeclarationFacts(computed.Properties.Any(pair =>
                !pair.Key.StartsWith("--", StringComparison.Ordinal) && ContainsCharacterUnit(pair.Value)),
            computed.HasCascadeStateMatching(IsLogicalPropertyName),
            computed.HasCascadeStateMatching(IsOptionalLayoutGroupProperty),
            computed.HasCascadeStateMatching(IsPaintProperty));
        if (_declarationFacts.Count < MaximumDeclarationFacts) _declarationFacts.Add(identity, facts);
        return facts;
    }

    private static bool IsOptionalLayoutGroupProperty(string name) =>
        string.Equals(name, "all", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "flex", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("flex-", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "columns", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("column-", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "grid", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("grid-", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "gap", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "row-gap", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("align-", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("justify-", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("place-", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "order", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "-webkit-box-orient", StringComparison.OrdinalIgnoreCase);

    private static bool IsPaintProperty(string name) =>
        string.Equals(name, "all", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "background", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("background-", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "border-image", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("border-image-", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "opacity", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "transform", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "scale", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "transform-origin", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "clip-path", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "box-decoration-break", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "box-shadow", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "text-shadow", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "-webkit-text-fill-color", StringComparison.OrdinalIgnoreCase)
        || UnimplementedProperties.Contains(name, StringComparer.OrdinalIgnoreCase);

    private readonly record struct DeclarationFacts(
        bool HasCharacterUnits, bool HasLogicalProperties, bool HasOptionalLayoutGroups, bool HasPaintState);
}
