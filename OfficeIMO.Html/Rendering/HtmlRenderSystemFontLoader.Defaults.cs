using AngleSharp.Dom;
using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class HtmlRenderSystemFontLoader {
    private readonly record struct DefaultFontVisibility(bool Blocked, bool PaintVisible, bool ZeroFont) {
        internal static DefaultFontVisibility Visible => new(false, true, false);
        internal bool CanPaint => !Blocked && PaintVisible && !ZeroFont;
    }

    /// <summary>Filters default host requests before loading, while allowing a visible child of hidden text.</summary>
    private static DefaultFontVisibility ResolveDefaultFontVisibility(IElement element, HtmlComputedStyle computed,
        DefaultFontVisibility parent, HtmlDisclosureState disclosures) {
        string tag = element.LocalName;
        bool blocked = parent.Blocked || disclosures.IsClosedChild(element)
            || tag is "head" or "style" or "script" or "template" or "noscript" or "meta" or "link" or "title" or "base"
            || tag == "input" && string.Equals(element.GetAttribute("type"), "hidden", StringComparison.OrdinalIgnoreCase)
            || HtmlRenderStyleResolver.ResolveDisplay(element, computed.GetValue("display")) == "none"
            || IsZeroOpacity(computed.GetValue("opacity"));
        string visibility = computed.GetValue("visibility").Trim().ToLowerInvariant();
        bool paintVisible = visibility is "" or "inherit" or "unset" ? parent.PaintVisible
            : visibility is not ("hidden" or "collapse");
        string fontSize = computed.GetValue("font-size").Trim();
        bool zeroFont = fontSize.Length == 0 || fontSize is "inherit" or "unset"
            ? parent.ZeroFont : IsAbsoluteZeroFontSize(fontSize);
        return new DefaultFontVisibility(blocked, paintVisible, zeroFont);
    }

    private static bool IsZeroOpacity(string value) {
        string number = value.Trim();
        if (number.EndsWith("%", StringComparison.Ordinal)) number = number.Substring(0, number.Length - 1);
        return double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double opacity)
            && opacity <= 0D;
    }

    // Missing relative bases deliberately fail rather than guessing that a CSS
    // expression which happens to be zero at one arbitrary size is always zero.
    private static bool IsAbsoluteZeroFontSize(string value) =>
        HtmlRenderCssValues.TryLength(value, double.NaN, double.NaN, double.NaN, out double size) && size == 0D;

    private static bool HasDirectDefaultFontText(IElement element, HtmlDisclosureState disclosures) {
        if (element.NamespaceUri != "http://www.w3.org/1999/xhtml") return false;
        if (element.ChildNodes.OfType<IText>().Any(text => !string.IsNullOrWhiteSpace(text.Data)
                && !disclosures.IsClosedChild(text))) return true;
        return element.LocalName == "input" && (!string.IsNullOrEmpty(element.GetAttribute("value"))
            || !string.IsNullOrEmpty(element.GetAttribute("placeholder"))
            || string.Equals(element.GetAttribute("type"), "submit", StringComparison.OrdinalIgnoreCase)
            || string.Equals(element.GetAttribute("type"), "reset", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Keeps a covered caller-owned CSS prefix from causing unused default host discovery.</summary>
    private static bool ScopedFacesCoverDirectText(IElement element, string families,
        OfficeFontFaceDescriptor descriptor, OfficeFontFaceCollection fonts, ISet<string> suppliedFamilies,
        HtmlDisclosureState disclosures,
        CancellationToken cancellationToken) {
        bool found = false;
        foreach (IText text in element.ChildNodes.OfType<IText>()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(text.Data) || disclosures.IsClosedChild(text)) continue;
            found = true;
            if (!Covers(text.Data)) return false;
        }
        if (element.LocalName == "input") {
            string? value = element.GetAttribute("value");
            if (string.IsNullOrEmpty(value)) value = element.GetAttribute("placeholder");
            if (!string.IsNullOrEmpty(value)) {
                found = true;
                if (!Covers(value!)) return false;
            }
        }
        return found;

        bool Covers(string value) {
            foreach (OfficeFontFallbackRun run in fonts.PlanFallbackRuns(value, families, descriptor)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!fonts.TryResolveFaceForText(run.Text, run.FamilyRequest, descriptor, out OfficeFontFace? face)
                    || face == null || !suppliedFamilies.Contains(face.FamilyName)) return false;
            }
            return true;
        }
    }

    private static bool CanPaintDefaultPseudo(HtmlComputedStyle pseudo, DefaultFontVisibility parent) {
        string visibility = pseudo.GetValue("visibility").Trim().ToLowerInvariant();
        bool paintVisible = visibility is "" or "inherit" or "unset" ? parent.PaintVisible
            : visibility is not ("hidden" or "collapse");
        string fontSize = pseudo.GetValue("font-size").Trim();
        bool zeroFont = fontSize.Length == 0 || fontSize is "inherit" or "unset"
            ? parent.ZeroFont : IsAbsoluteZeroFontSize(fontSize);
        return !parent.Blocked && paintVisible && !zeroFont
            && !pseudo.GetValue("display").Trim().Equals("none", StringComparison.OrdinalIgnoreCase)
            && !IsZeroOpacity(pseudo.GetValue("opacity"));
    }

    private static bool TryGetDefaultPseudoRequest(IElement element, HtmlPseudoElementKind kind,
        HtmlComputedStyle pseudo, out bool needsText) {
        needsText = kind is HtmlPseudoElementKind.FirstLine or HtmlPseudoElementKind.FirstLetter or HtmlPseudoElementKind.Marker;
        if (needsText) return true;
        if (kind == HtmlPseudoElementKind.Placeholder)
            return string.IsNullOrEmpty(element.GetAttribute("value")) && !string.IsNullOrEmpty(element.GetAttribute("placeholder"));
        string content = pseudo.GetValue("content").Trim();
        return content.Length > 0 && !content.Equals("none", StringComparison.OrdinalIgnoreCase)
            && !content.Equals("normal", StringComparison.OrdinalIgnoreCase)
            && content is not ("''" or "\"\"");
    }
}
