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
            computed.HasCascadeStateMatching(IsLogicalPropertyName));
        if (_declarationFacts.Count < MaximumDeclarationFacts) _declarationFacts.Add(identity, facts);
        return facts;
    }

    private readonly record struct DeclarationFacts(bool HasCharacterUnits, bool HasLogicalProperties);
}
