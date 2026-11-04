using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    /// <summary>Returns the footnote-call font when CSS does not override its user-agent default.</summary>
    internal static double ResolveDefaultFootnoteCallFontSize(double parentFontSize) => parentFontSize * 0.7D;

    private double ResolveComputedFontSize(
        string tag,
        HtmlComputedStyle computed,
        string fontSizeValue,
        double parentFontSize,
        bool pseudoElement,
        string pseudoSemanticRole) {
        // Call defaults must precede line-height and other font-relative properties.
        // Explicit inherit/reset declarations retain the ordinary CSS font contract.
        if (pseudoElement && pseudoSemanticRole == "footnote-call" && !HasAuthoredValue(computed, "font-size")) {
            return ResolveDefaultFootnoteCallFontSize(parentFontSize);
        }
        double defaultFontSize = pseudoElement ? parentFontSize : ResolveDefaultTagFontSize(tag, parentFontSize);
        return computed.IsOriginRevertedValue("font-size")
            ? defaultFontSize
            : computed.IsResetValue("font-size")
            ? _options.DefaultFontSize
            : computed.IsImplicitlyInheritedValue("font-size")
            ? defaultFontSize
            : computed.IsInheritedValue("font-size")
            ? parentFontSize
            : string.IsNullOrWhiteSpace(fontSizeValue)
            ? defaultFontSize
            : ResolveFontSize(fontSizeValue, parentFontSize,
                tag == "html" && !pseudoElement ? _options.DefaultFontSize : _rootFontSize);
    }

    /// <summary>Resolves Browser sizing with the same provenance policy used by computed container values.</summary>
    private HtmlFontSizeContext ResolveBrowserFontSize(
        string tag,
        HtmlComputedStyle computed,
        string family,
        HtmlRenderBoxStyle? parent,
        bool pseudoElement,
        string pseudoSemanticRole) {
        HtmlRenderingFontDefaults defaults = _fontDefaults!.Value;
        HtmlFontSizeContext inherited = parent?.FontSizeContext
            ?? (parent == null ? defaults.Initial : new HtmlFontSizeContext(parent.Font.Size, true, string.Empty, false));
        HtmlFontSizeContext fallback = pseudoElement ? inherited : HtmlFontSizePolicy.ResolveHtmlDefault(tag, inherited, browser: true);
        if (pseudoElement && pseudoSemanticRole == "footnote-call" && !HasAuthoredValue(computed, "font-size")) {
            fallback = inherited.WithRelativeSize(ResolveDefaultFootnoteCallFontSize(inherited.Pixels));
        }
        return HtmlFontSizePolicy.Resolve(computed, inherited, family, defaults, fallback,
            new OfficeIMO.Html.Css.HtmlCssLengthResolutionContext {
                PercentageReference = inherited.Pixels,
                FontSize = inherited.Pixels,
                RootFontSize = tag == "html" && !pseudoElement ? defaults.Initial.Pixels : _rootFontSize,
                ViewportWidth = _viewportWidth,
                ViewportHeight = _viewportHeight,
                ContainerWidth = double.IsNaN(_activeContainerWidth) ? null : _activeContainerWidth,
                ContainerHeight = double.IsNaN(_activeContainerHeight) ? null : _activeContainerHeight,
                ContainerInlineSize = double.IsNaN(_activeContainerWidth) ? null : _activeContainerWidth,
                ContainerBlockSize = double.IsNaN(_activeContainerHeight) ? null : _activeContainerHeight,
                CharacterAdvance = double.IsNaN(_activeCharacterAdvance) ? null : _activeCharacterAdvance,
                UprightVerticalText = _activeUprightVerticalText
            });
    }

    internal static OfficeFontFaceDescriptor ResolveFontFaceDescriptor(
        string tag,
        HtmlComputedStyle computed,
        OfficeFontFaceDescriptor inherited) {
        // HTML defaults are materialized by the computed-style owner before
        // inheritance, including the normal values for explicit initial resets.
        string weightValue = computed.IsImplicitlyInheritedValue("font-weight")
            ? string.Empty
            : computed.GetValue("font-weight");
        if (tag == "math" && (!HasAuthoredValue(computed, "font-weight") || computed.IsResetValue("font-weight")))
            weightValue = "normal";
        int weight = OfficeFontFaceCssParser.TryWeight(weightValue, inherited.Weight, out int parsedWeight) ? parsedWeight : inherited.Weight;
        double stretch = OfficeFontFaceCssParser.TryStretch(computed.GetValue("font-stretch"), inherited.StretchPercent, out double parsedStretch)
            ? parsedStretch : inherited.StretchPercent;
        string slantValue = tag == "math" && (!HasAuthoredValue(computed, "font-style") || computed.IsResetValue("font-style"))
            ? "normal" : computed.GetValue("font-style");
        if (!OfficeFontFaceCssParser.TrySlant(slantValue, inherited, out OfficeFontSlant slant, out double angle)) {
            slant = inherited.Slant;
            angle = inherited.ObliqueAngleDegrees;
        }
        return new OfficeFontFaceDescriptor(weight, stretch, slant, angle);
    }

    /// <summary>Preserves the user-agent monospace default of code elements when a font family is reset.</summary>
    internal static string ResolveInitialFontFamily(string tag, HtmlRenderOptions options) =>
        UsesDocumentMonospaceFamily(tag, options)
            ? options.UserAgentStyles == HtmlRenderUserAgentStyleMode.Document ? "Consolas" : "monospace"
            : options.DefaultFontFamily;

    /// <summary>
    /// Code elements carry a user-agent <c>font-family: monospace</c> rule in both style modes. Like any
    /// declared value it wins over inheritance, so a proportional family set on an ancestor never reaches code.
    /// </summary>
    internal static bool UsesDocumentMonospaceFamily(string tag, HtmlRenderOptions options) =>
        tag is "code" or "pre" or "kbd" or "samp" or "tt" or "listing" or "xmp" or "plaintext";

    /// <summary>Resolves an element family, letting the user-agent monospace rule beat implicit inheritance.</summary>
    internal static string ResolveElementFontFamily(string tag, HtmlComputedStyle computed, string fallback, HtmlRenderOptions options) {
        string initial = ResolveInitialFontFamily(tag, options);
        if (UsesDocumentMonospaceFamily(tag, options)
            && (!HasAuthoredValue(computed, "font-family") || computed.IsImplicitlyInheritedValue("font-family"))) {
            return initial;
        }

        return ResolveFontFamily(tag, computed, fallback, initial);
    }

    internal static string ResolveFontFamily(string tag, HtmlComputedStyle computed, string inherited, string? initialFamily = null) {
        if (tag == "math") {
            if (!HasAuthoredValue(computed, "font-family")) return "math";
            if (computed.IsResetValue("font-family"))
                return HtmlRenderCssValues.FontFamilyList(computed.GetValue("font-family"), "serif");
            if (computed.IsInheritedValue("font-family")) return inherited;
        }
        string fallback = computed.IsResetValue("font-family") && !computed.IsOriginRevertedValue("font-family")
            ? initialFamily ?? inherited : inherited;
        return HtmlRenderCssValues.FontFamilyList(computed.GetValue("font-family"), fallback);
    }

}
