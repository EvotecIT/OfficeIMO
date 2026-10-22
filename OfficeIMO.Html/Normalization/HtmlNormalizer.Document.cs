using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlNormalizer {
    // Correlation belongs outside HTML: only actual source/normalized markup
    // consumes its character boundary, and authored attributes remain untouched.
    internal static IHtmlDocument NormalizeToDocument(IHtmlDocument source, HtmlNormalizationOptions options, CancellationToken cancellationToken = default,
        Action<IElement, IElement>? elementNormalized = null) {
        cancellationToken.ThrowIfCancellationRequested();
        HtmlConversionInputGuard.ValidateDocument(source, options.Limits, cancellationToken);
        bool hasFormState = source.QuerySelectorAll("input,textarea,select,option")
            .Any(element => NativeFormState.Get(element) != null);
        if (!hasFormState && elementNormalized == null)
            return HtmlDocumentParser.ParseDocument(Normalize(source, options, cancellationToken), cancellationToken);

        var emittedElements = new Dictionary<int, IElement>();
        string markup = Normalize(source, options, cancellationToken, emittedElements);
        IHtmlDocument normalized = HtmlDocumentParser.ParseDocument(markup, cancellationToken);
        foreach (IElement element in normalized.QuerySelectorAll("*")) {
            cancellationToken.ThrowIfCancellationRequested();
            int? offset = element.SourceReference?.Position.Index;
            // Parser-created wrappers and omitted nodes have no matching emitted
            // tag. Never guess by ordinal or transfer identity across namespaces.
            if (!offset.HasValue || !emittedElements.TryGetValue(offset.Value, out IElement? original)
                || !string.Equals(element.LocalName, original.LocalName, StringComparison.Ordinal)
                || !string.Equals(element.NamespaceUri, original.NamespaceUri, StringComparison.Ordinal)) continue;
            NativeFormState.Attach(element, NativeFormState.Get(original));
            elementNormalized?.Invoke(original, element);
        }
        NativeFormState.ApplyTree(normalized, cancellationToken);
        // The operation-local map is released here. Only immutable live state
        // and caller-selected side metadata are attached to returned elements.
        return normalized;
    }
}
