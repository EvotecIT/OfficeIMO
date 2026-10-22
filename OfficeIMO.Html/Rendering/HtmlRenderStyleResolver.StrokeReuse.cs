using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private const int MaximumReusedBoxStrokes = 1024;
    private readonly Dictionary<BoxStrokeContext, ResolvedBoxStroke> _boxStrokes = new();

    /// <summary>Reuses used border/outline values for one immutable computed style and the exact numeric context.</summary>
    /// <remarks>Layout callers receive value copies; parent boxes and other mutable resolved style fields are never cached.</remarks>
    private ResolvedBoxStroke ResolveBoxStroke(HtmlComputedStyle computed, double reference, double fontSize,
        OfficeColor color) {
        var context = new BoxStrokeContext(computed, reference, fontSize, _rootFontSize,
            _viewportWidth, _viewportHeight, color, _activeCharacterAdvance, _activeUprightVerticalText);
        if (_boxStrokes.TryGetValue(context, out ResolvedBoxStroke found)) return found;

        bool borderSupported = HtmlCssBoxStrokeParser.TryParseBorder(computed, reference, fontSize,
            _rootFontSize, _viewportWidth, _viewportHeight, color, out HtmlRenderBorderEdges borders,
            out string borderDetail, _activeCharacterAdvance, _activeUprightVerticalText);
        bool outlineSupported = HtmlCssBoxStrokeParser.TryParseOutline(computed, reference, fontSize,
            _rootFontSize, _viewportWidth, _viewportHeight, color, out double outlineWidth,
            out string outlineStyle, out OfficeColor outlineColor, out bool outlineInvert,
            out double outlineOffset, out string outlineDetail, _activeCharacterAdvance, _activeUprightVerticalText);
        var result = new ResolvedBoxStroke(HtmlCssBoxStrokeParser.HasBorderDeclaration(computed),
            borderSupported, borders, borderDetail, outlineSupported, outlineWidth, outlineStyle,
            outlineColor, outlineInvert, outlineOffset, outlineDetail);
        if (_boxStrokes.Count < MaximumReusedBoxStrokes) _boxStrokes.Add(context, result);
        return result;
    }

    private readonly record struct BoxStrokeContext(HtmlComputedStyle Computed, double Reference,
        double FontSize, double RootFontSize, double ViewportWidth, double ViewportHeight,
        OfficeColor Color, double CharacterAdvance, bool UprightVerticalText);

    private readonly record struct ResolvedBoxStroke(bool BorderDeclared, bool BorderSupported,
        HtmlRenderBorderEdges Borders, string BorderDetail, bool OutlineSupported, double OutlineWidth,
        string OutlineStyle, OfficeColor OutlineColor, bool OutlineColorInvert, double OutlineOffset,
        string OutlineDetail);
}
