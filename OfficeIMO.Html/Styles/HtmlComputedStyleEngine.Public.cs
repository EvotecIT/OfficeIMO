using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>Parses bounded HTML and computes styles for matching elements.</summary>
    public static IReadOnlyDictionary<Dom.HtmlElement, HtmlComputedStyle> Compute(
        string html,
        HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) =>
        Compute(HtmlConversionDocument.Parse(html), mediaContext);

    /// <summary>Parses bounded HTML and computes styles with explicit inspection options.</summary>
    public static IReadOnlyDictionary<Dom.HtmlElement, HtmlComputedStyle> Compute(
        string html,
        HtmlComputedStyleOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        return Compute(HtmlConversionDocument.Parse(html), options);
    }

    /// <summary>Computes styles from a retained conversion document without reparsing its HTML source.</summary>
    public static IReadOnlyDictionary<Dom.HtmlElement, HtmlComputedStyle> Compute(
        HtmlConversionDocument document,
        HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return Compute(document.Document, mediaContext, document.Limits);
    }

    /// <summary>Computes styles from a retained conversion document with explicit inspection options.</summary>
    public static IReadOnlyDictionary<Dom.HtmlElement, HtmlComputedStyle> Compute(
        HtmlConversionDocument document,
        HtmlComputedStyleOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (options == null) throw new ArgumentNullException(nameof(options));
        return Compute(document.Document, options, document.Limits);
    }

    /// <summary>Computes styles keyed by owned nodes from the supplied document snapshot.</summary>
    /// <remarks>A prepared document retains the unbounded computation contract. To apply input and CSS
    /// budgets, create an HtmlConversionDocument with explicit limits and use that overload.</remarks>
    public static IReadOnlyDictionary<Dom.HtmlElement, HtmlComputedStyle> Compute(
        Dom.HtmlDocument document,
        HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) =>
        Compute(document, mediaContext, limits: null);

    /// <summary>Computes styles keyed by owned nodes with explicit inspection options.</summary>
    public static IReadOnlyDictionary<Dom.HtmlElement, HtmlComputedStyle> Compute(
        Dom.HtmlDocument document,
        HtmlComputedStyleOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        return Compute(document, options, limits: null);
    }

    private static IReadOnlyDictionary<Dom.HtmlElement, HtmlComputedStyle> Compute(
        Dom.HtmlDocument document,
        HtmlCssMediaContext mediaContext,
        HtmlConversionLimits? limits) =>
        Compute(document, new HtmlComputedStyleOptions { MediaContext = mediaContext }, limits);

    private static IReadOnlyDictionary<Dom.HtmlElement, HtmlComputedStyle> Compute(
        Dom.HtmlDocument document,
        HtmlComputedStyleOptions options,
        HtmlConversionLimits? limits) {
        IHtmlDocument native = NativeDomBridge.GetNativeDocument(document);
        if (limits != null) HtmlConversionInputGuard.ValidateDocument(native, limits);
        var state = NativeDomBridge.GetState(document);
        HtmlComputedStyleSet computed = ComputeStyleSet(native, MediaEnvironment.CreateDefault(options.MediaContext),
            false, options.IncludeCascadeTraces, limits);
        return computed.Elements.ToDictionary(pair => (Dom.HtmlElement)state.ToOwned[pair.Key], pair => pair.Value);
    }
}
