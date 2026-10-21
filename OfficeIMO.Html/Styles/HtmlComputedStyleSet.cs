using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal enum HtmlPseudoElementKind {
    Before,
    After,
    Marker,
    FootnoteCall,
    FootnoteMarker,
    FirstLetter,
    FirstLine,
    Placeholder
}

internal sealed class HtmlComputedStyleSet {
    private readonly IReadOnlyDictionary<IElement, HtmlPseudoElementStylePair> _pseudoElements;

    internal HtmlComputedStyleSet(
        IReadOnlyDictionary<IElement, HtmlComputedStyle> elements,
        IReadOnlyDictionary<IElement, HtmlPseudoElementStylePair> pseudoElements,
        HtmlRenderingFontDefaults? fontDefaults = null,
        IReadOnlyDictionary<string, HtmlComputedStyleEngine.CustomPropertyRegistration>? registrations = null,
        HtmlCssProcessingBudget? budget = null) {
        Elements = elements ?? throw new ArgumentNullException(nameof(elements));
        FontDefaults = fontDefaults;
        _registrations = registrations;
        _budget = budget;
        _pseudoElements = pseudoElements ?? throw new ArgumentNullException(nameof(pseudoElements));
    }

    internal IReadOnlyDictionary<IElement, HtmlComputedStyle> Elements { get; }
    internal HtmlRenderingFontDefaults? FontDefaults { get; }

    private readonly IReadOnlyDictionary<string, HtmlComputedStyleEngine.CustomPropertyRegistration>? _registrations;
    private readonly HtmlCssProcessingBudget? _budget;
    private long _generatedContextApplications;

    /// <summary>Preserves registered custom-property semantics and their operation-wide fanout limit for paged contexts.</summary>
    internal HtmlComputedStyle ComputeGeneratedContext(HtmlComputedStyle parent,
        IEnumerable<HtmlGeneratedStyleDeclaration> declarations, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_registrations?.Count > 0)
            _budget?.ValidateRegistrationFanout(_registrations.Count, Elements.Count + ++_generatedContextApplications);
        return HtmlComputedStyleEngine.ComputeGeneratedContext(parent, declarations, _budget?.HasDeclarationLimit == true,
            cancellationToken, FontDefaults, _registrations);
    }

    internal bool HasPseudoElements => _pseudoElements.Count > 0;

    internal bool TryGetPseudoStyle(IElement element, HtmlPseudoElementKind kind, out HtmlComputedStyle style) {
        if (_pseudoElements.TryGetValue(element, out HtmlPseudoElementStylePair? pair)) {
            HtmlComputedStyle? found = kind switch {
                HtmlPseudoElementKind.Before => pair.Before,
                HtmlPseudoElementKind.After => pair.After,
                HtmlPseudoElementKind.Marker => pair.Marker,
                HtmlPseudoElementKind.FootnoteCall => pair.FootnoteCall,
                HtmlPseudoElementKind.FootnoteMarker => pair.FootnoteMarker,
                HtmlPseudoElementKind.FirstLetter => pair.FirstLetter,
                HtmlPseudoElementKind.FirstLine => pair.FirstLine,
                _ => pair.Placeholder
            };
            if (found != null) {
                style = found;
                return true;
            }
        }

        style = null!;
        return false;
    }
}

internal sealed class HtmlPseudoElementStylePair {
    internal HtmlComputedStyle? Before { get; set; }
    internal HtmlComputedStyle? After { get; set; }
    internal HtmlComputedStyle? Marker { get; set; }
    internal HtmlComputedStyle? FootnoteCall { get; set; }
    internal HtmlComputedStyle? FootnoteMarker { get; set; }
    internal HtmlComputedStyle? FirstLetter { get; set; }
    internal HtmlComputedStyle? FirstLine { get; set; }
    internal HtmlComputedStyle? Placeholder { get; set; }
}
