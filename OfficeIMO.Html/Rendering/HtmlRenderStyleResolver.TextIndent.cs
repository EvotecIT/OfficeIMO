using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HashSet<IElement> _reportedTextIndentFallbacks = new HashSet<IElement>();

    private HtmlRenderTextIndent? ResolveTextIndent(IElement element, HtmlComputedStyle computed,
        double fontSize, HtmlRenderBoxStyle? parent) {
        if (computed.IsResetValue("text-indent")) return null;
        string value = computed.GetValue("text-indent").Trim().ToLowerInvariant();
        if (computed.IsInheritedValue("text-indent") || value.Length == 0 || value == "inherit" || value == "unset") {
            return parent?.TextIndent;
        }
        if (value == "initial") return null;
        var indent = new HtmlRenderTextIndent(value, fontSize, _options.DefaultFontSize,
            _viewportWidth, _viewportHeight, _activeContainerWidth, _activeContainerHeight, _activeCharacterAdvance);
        if (HtmlRenderCssValues.HasExplicitLengthSyntax(value, allowPercentage: true, allowUnitlessZero: true)
            && indent.TryResolve(100D, out _)) return indent;
        if (_reportedTextIndentFallbacks.Add(element)) {
            _diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.TextIndentValueUnsupported,
                "A text-indent value could not be rendered; the inherited indentation was used.",
                HtmlDiagnosticSeverity.Warning, DescribeSource(element), "text-indent=" + value);
        }
        return parent?.TextIndent;
    }
}

/// <summary>
/// Keeps the declaring element's length context through inheritance. Percentages
/// are resolved against the formatted block's inline size, after its box sizing.
/// </summary>
internal sealed class HtmlRenderTextIndent {
    private readonly string _value;
    private readonly double _fontSize, _rootFontSize, _viewportWidth, _viewportHeight, _containerWidth, _containerHeight, _characterAdvance;

    internal HtmlRenderTextIndent(string value, double fontSize, double rootFontSize,
        double viewportWidth, double viewportHeight, double containerWidth, double containerHeight, double characterAdvance) {
        _value = value;
        _fontSize = fontSize;
        _rootFontSize = rootFontSize;
        _viewportWidth = viewportWidth;
        _viewportHeight = viewportHeight;
        _containerWidth = containerWidth;
        _containerHeight = containerHeight;
        _characterAdvance = characterAdvance;
    }

    internal bool TryResolve(double width, out double result) => HtmlRenderCssValues.TryLength(
        _value, width, _fontSize, _rootFontSize, _viewportWidth, _viewportHeight,
        _containerWidth, _containerHeight, out result, _characterAdvance);

    internal double Resolve(double width) => TryResolve(width, out double result) ? result : 0D;
}
