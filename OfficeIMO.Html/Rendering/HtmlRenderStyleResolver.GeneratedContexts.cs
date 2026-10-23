using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    /// <summary>Supplies the document-root parent after the final rendering cascade, never the body's styles.</summary>
    internal (HtmlComputedStyle Computed, HtmlRenderBoxStyle Used) ResolvePageStyleParent() {
        foreach (var pair in _computedStyles.Elements) {
            if (pair.Key.LocalName.Equals("html", StringComparison.OrdinalIgnoreCase))
                return (pair.Value, Resolve(pair.Key, _viewportWidth));
        }
        return (new HtmlComputedStyle(new Dictionary<string, string>()), new HtmlRenderBoxStyle {
            Font = new OfficeFontInfo(_options.DefaultFontFamily, _options.DefaultFontSize),
            LineHeight = _options.DefaultFontSize * _options.DefaultLineHeight,
            Color = OfficeColor.Black
        });
    }

    internal HtmlComputedStyleSet PageComputedStyles => _computedStyles;
    internal System.Threading.CancellationToken PageCancellationToken => _cancellationToken;

    /// <summary>Uses the ordinary font sizing and leading owners for a computed non-DOM paged context.</summary>
    internal HtmlRenderBoxStyle ResolvePageTypography(
        HtmlComputedStyle computed, HtmlRenderBoxStyle parent, double viewportWidth, double viewportHeight) {
        double previousWidth = _activeContainerWidth, previousHeight = _activeContainerHeight;
        double previousViewportWidth = _viewportWidth, previousViewportHeight = _viewportHeight;
        double previousAdvance = _activeCharacterAdvance;
        bool previousUpright = _activeUprightVerticalText;
        try {
            _viewportWidth = _activeContainerWidth = viewportWidth;
            _viewportHeight = _activeContainerHeight = viewportHeight;
            _activeUprightVerticalText = false;
            _activeCharacterAdvance = ContainsCharacterUnit(computed.GetValue("font-size"))
                ? _measureCharacterAdvance(parent) : double.NaN;
            string family = ResolveFontFamily(string.Empty, computed, parent.Font.FamilyName, _options.DefaultFontFamily);
            HtmlFontSizeContext? context = _fontDefaults.HasValue
                ? ResolveBrowserFontSize(string.Empty, computed, family, parent, true, string.Empty) : null;
            double size = context?.Pixels ?? ResolveComputedFontSize(string.Empty, computed,
                computed.GetValue("font-size"), parent.Font.Size, true, string.Empty);
            OfficeFontFaceDescriptor descriptor = ResolveFontFaceDescriptor(string.Empty, computed, parent.FontDescriptor);
            var font = new OfficeFontInfo(family, size, descriptor, descriptor.ToStyle());
            _activeCharacterAdvance = ContainsCharacterUnit(computed.GetValue("line-height"))
                ? _measureCharacterAdvance(new HtmlRenderBoxStyle { Font = font, FontDescriptor = descriptor }) : double.NaN;
            double leading = ResolveComputedLineHeight(computed, size, parent, out double alignment, out bool natural);
            OfficeColor fallback = computed.IsResetValue("color") ? OfficeColor.Black : parent.Color;
            OfficeColor color = HtmlRenderCssValues.TryColor(computed.GetValue("color"), out OfficeColor parsed) ? parsed : fallback;
            TryResolveBackgroundColor(computed, color, out OfficeColor? background);
            return new HtmlRenderBoxStyle {
                BackgroundColor = background,
                Font = font, FontDescriptor = descriptor, FontSizeContext = context, Color = color,
                LineHeight = leading, InlineAlignmentLineHeight = alignment, UsesNaturalLineHeight = natural
            };
        } finally {
            _viewportWidth = previousViewportWidth;
            _viewportHeight = previousViewportHeight;
            _activeContainerWidth = previousWidth;
            _activeContainerHeight = previousHeight;
            _activeCharacterAdvance = previousAdvance;
            _activeUprightVerticalText = previousUpright;
        }
    }
}
