namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>Reuses only context-free declaration grammar and shorthand expansion within one computation.</summary>
    /// <remarks>Selector/declaration budgets, CSS-wide inheritance, variables, cascade candidates and provenance still execute per element.</remarks>
    private sealed class DeclarationSyntaxReuseCache {
        private const int MaximumEntries = 256;
        private const int MaximumPropertyLength = 128;
        private const int MaximumValueLength = 256;
        private const int MaximumLonghands = 64;
        private readonly Dictionary<(string Name, string Value), bool> _supported = new();
        private readonly Dictionary<(string Name, string Value), Expansion> _expansions = new();

        internal bool IsSupported(string name, string value) {
            if (!Eligible(name, value)) return IsSupportedDeclarationValue(name, value);
            var key = (name, value);
            if (_supported.TryGetValue(key, out bool found)) return found;
            bool supported = IsSupportedDeclarationValue(name, value);
            if (_supported.Count < MaximumEntries) _supported.Add(key, supported);
            return supported;
        }

        internal bool TryExpand(string name, string value, out IReadOnlyList<KeyValuePair<string, string>> longhands) {
            if (!Eligible(name, value)) return TryExpandCascadeShorthand(name, value, out longhands);
            var key = (name, value);
            if (_expansions.TryGetValue(key, out Expansion found)) {
                longhands = found.Longhands;
                return found.Expanded;
            }
            bool expanded = TryExpandCascadeShorthand(name, value, out longhands);
            if (_expansions.Count < MaximumEntries && longhands.Count <= MaximumLonghands) {
                longhands = longhands.ToArray();
                _expansions.Add(key, new Expansion(expanded, longhands));
            }
            return expanded;
        }

        private static bool Eligible(string name, string value) =>
            name.Length <= MaximumPropertyLength && value.Length <= MaximumValueLength;

        private readonly record struct Expansion(bool Expanded, IReadOnlyList<KeyValuePair<string, string>> Longhands);
    }

    private static bool IsSupportedDeclarationValueWithReuse(string name, string value, DeclarationSyntaxReuseCache? reuse) =>
        reuse?.IsSupported(name, value) ?? IsSupportedDeclarationValue(name, value);

    private static bool TryExpandCascadeShorthandWithReuse(string name, string value,
        out IReadOnlyList<KeyValuePair<string, string>> longhands, DeclarationSyntaxReuseCache? reuse) =>
        reuse == null ? TryExpandCascadeShorthand(name, value, out longhands) : reuse.TryExpand(name, value, out longhands);
}
