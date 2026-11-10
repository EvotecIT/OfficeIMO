using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>
    /// Shares one cascade among elements that selectors cannot tell apart, as browsers do for siblings.
    /// </summary>
    /// <remarks>
    /// Each element gets a token for its namespace, tag and complete attribute list together with its parent's
    /// token, so equal tokens mean the same tags and attributes on every ancestor. Without sibling, position or
    /// state selectors, every rule then matches elements with equal tokens alike, and the remaining cascade inputs
    /// (parent style, container context and geometry) complete the key. Elements whose identity, inline style or
    /// text can change their style keep their own cascade. Selector matching still runs for a shared element, so
    /// the processing budget is charged exactly as without sharing; only the cascade itself is reused.
    /// </remarks>
    private sealed class StyleSharingCache {
        private const int MaximumEntries = 4096;
        private const int MaximumAttributeCharacters = 1024;
        private readonly bool _enabled;
        private readonly Dictionary<IElement, object> _tokens = new();
        private readonly Dictionary<(object Parent, string LocalName, string Attributes), object> _internedTokens = new();
        private readonly Dictionary<SharedStyleKey, HtmlComputedStyle> _styles = new();

        internal StyleSharingCache(bool enabled) => _enabled = enabled;

        /// <summary>Rejects selectors that read siblings, position, state or other non-attribute facts.</summary>
        internal static bool IsPositionIndependent(string selector) =>
            selector.IndexOf(':') < 0 && selector.IndexOf('+') < 0 && selector.IndexOf('~') < 0;

        internal SharedStyleKey? CreateKey(IElement element, HtmlComputedStyle? parent, bool includeCascadeTraces,
            IReadOnlyList<ContainerQueryContext> containerContexts, double containingWidth, double? containingHeight) {
            if (!_enabled || includeCascadeTraces) return null;
            object token = CreateToken(element);
            if (token is not SharedToken) return null;
            return new SharedStyleKey(token, parent?.DeclarationIdentity, containerContexts, containingWidth, containingHeight);
        }

        /// <summary>Gets the cascade result, before per-element sizing, of an element with the same key.</summary>
        internal bool TryGet(SharedStyleKey key, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out HtmlComputedStyle? template) =>
            _styles.TryGetValue(key, out template);

        internal void Add(SharedStyleKey key, HtmlComputedStyle template) {
            if (_styles.Count < MaximumEntries) _styles[key] = template;
        }

        private object CreateToken(IElement element) {
            object? parentToken = element.ParentElement is { } parentElement && _tokens.TryGetValue(parentElement, out object? found)
                ? found : null;
            object token;
            string? attributes = parentToken == null || !IsShareable(element) ? null : DescribeAttributes(element);
            if (attributes == null) {
                token = new object();
            } else {
                var internKey = (parentToken!, element.LocalName, attributes);
                if (!_internedTokens.TryGetValue(internKey, out token!)) {
                    token = new SharedToken();
                    if (_internedTokens.Count < MaximumEntries) _internedTokens.Add(internKey, token);
                }
            }
            _tokens[element] = token;
            return token;
        }

        // Identity, inline declarations and text-derived direction belong to one element only.
        private static bool IsShareable(IElement element) =>
            string.Equals(element.NamespaceUri, "http://www.w3.org/1999/xhtml", StringComparison.Ordinal)
            && !element.HasAttribute("id") && !element.HasAttribute("style")
            && !string.Equals(element.LocalName, "bdi", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(element.GetAttribute("dir")?.Trim(), "auto", StringComparison.OrdinalIgnoreCase);

        private static string? DescribeAttributes(IElement element) {
            INamedNodeMap attributes = element.Attributes;
            if (attributes.Length == 0) return string.Empty;
            var text = new System.Text.StringBuilder();
            foreach (IAttr attribute in attributes) {
                text.Append(attribute.NamespaceUri).Append('\u0001').Append(attribute.Name).Append('\u0001')
                    .Append(attribute.Value).Append('\u0002');
                if (text.Length > MaximumAttributeCharacters) return null;
            }
            return text.ToString();
        }

        private sealed class SharedToken { }
    }

    private sealed class SharedStyleKey : IEquatable<SharedStyleKey> {
        private readonly object _token;
        private readonly object? _parentDeclarations;
        private readonly IReadOnlyList<ContainerQueryContext> _containerContexts;
        private readonly double _containingWidth;
        private readonly double? _containingHeight;

        internal SharedStyleKey(object token, object? parentDeclarations, IReadOnlyList<ContainerQueryContext> containerContexts,
            double containingWidth, double? containingHeight) {
            _token = token;
            _parentDeclarations = parentDeclarations;
            _containerContexts = containerContexts;
            _containingWidth = containingWidth;
            _containingHeight = containingHeight;
        }

        public bool Equals(SharedStyleKey? other) => other != null
            && ReferenceEquals(_token, other._token)
            && ReferenceEquals(_parentDeclarations, other._parentDeclarations)
            && ReferenceEquals(_containerContexts, other._containerContexts)
            && _containingWidth.Equals(other._containingWidth)
            && Nullable.Equals(_containingHeight, other._containingHeight);

        public override bool Equals(object? obj) => Equals(obj as SharedStyleKey);

        public override int GetHashCode() {
            unchecked {
                int hash = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_token);
                hash = hash * 31 + (_parentDeclarations == null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_parentDeclarations));
                hash = hash * 31 + System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_containerContexts);
                hash = hash * 31 + _containingWidth.GetHashCode();
                return hash * 31 + _containingHeight.GetHashCode();
            }
        }
    }
}
