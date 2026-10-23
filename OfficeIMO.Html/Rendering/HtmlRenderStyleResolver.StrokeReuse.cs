using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private const int MaximumReusedBoxStrokes = 1024;
    private readonly Dictionary<BoxStrokeContext, ResolvedBoxStroke> _boxStrokes = new();
    private readonly Dictionary<object, HtmlComputedStyle> _strokeStyleSources = new();
    private readonly Dictionary<HtmlCssBoxStrokeParser.StrokeDeclarations, HtmlComputedStyle> _strokeStyles = new();

    /// <summary>Reuses used border/outline values for one immutable computed style and the exact numeric context.</summary>
    /// <remarks>Layout callers receive value copies; parent boxes and other mutable resolved style fields are never cached.</remarks>
    private ResolvedBoxStroke ResolveBoxStroke(HtmlComputedStyle computed, double reference, double fontSize,
        OfficeColor color) {
        computed = ReuseStrokeDeclarations(computed);
        bool reuse = _boxStrokes.Count < MaximumReusedBoxStrokes;
        BoxStrokeContext context = default;
        if (reuse) {
            context = new BoxStrokeContext(computed, reference, fontSize, _rootFontSize,
                _viewportWidth, _viewportHeight, color, _activeCharacterAdvance, _activeUprightVerticalText);
            if (_boxStrokes.TryGetValue(context, out ResolvedBoxStroke found)) return found;
        }

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
        if (reuse) _boxStrokes.Add(context, result);
        return result;
    }

    /// <summary>Shares parser input only when every relevant immutable declaration and cascade state matches.</summary>
    /// <remarks>Source styles and mutable boxes remain distinct; only the existing numeric-context cache uses the representative.</remarks>
    private HtmlComputedStyle ReuseStrokeDeclarations(HtmlComputedStyle computed) {
        if (_strokeStyleSources.TryGetValue(computed.DeclarationIdentity, out HtmlComputedStyle? source)) return source;
        if (_strokeStyleSources.Count >= MaximumReusedBoxStrokes) return computed;

        var declarations = new HtmlCssBoxStrokeParser.StrokeDeclarations(computed);
        if (!_strokeStyles.TryGetValue(declarations, out source)) {
            source = computed;
            _strokeStyles.Add(declarations, source);
        }
        _strokeStyleSources.Add(computed.DeclarationIdentity, source);
        return source;
    }

    private readonly record struct BoxStrokeContext(HtmlComputedStyle Computed, double Reference,
        double FontSize, double RootFontSize, double ViewportWidth, double ViewportHeight,
        OfficeColor Color, double CharacterAdvance, bool UprightVerticalText);

    private readonly record struct ResolvedBoxStroke(bool BorderDeclared, bool BorderSupported,
        HtmlRenderBorderEdges Borders, string BorderDetail, bool OutlineSupported, double OutlineWidth,
        string OutlineStyle, OfficeColor OutlineColor, bool OutlineColorInvert, double OutlineOffset,
        string OutlineDetail);
}
