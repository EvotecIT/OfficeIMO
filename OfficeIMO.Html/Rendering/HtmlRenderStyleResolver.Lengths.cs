namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HtmlRenderLengthExpressionCache _lengthExpressions = new();

    private bool TryResolveLength(string? value, double reference, double fontSize, double rootFontSize, out double result) =>
        TryResolveLength(value, reference, fontSize, rootFontSize, out result, out _);

    private bool TryResolveLength(
        string? value,
        double reference,
        double fontSize,
        double rootFontSize,
        out double result,
        out bool isCalculated) =>
        HtmlRenderCssValues.TryLength(
            value,
            reference,
            fontSize,
            rootFontSize,
            _viewportWidth,
            _viewportHeight,
            _activeContainerWidth,
            _activeContainerHeight,
            out result,
            out isCalculated,
            _activeCharacterAdvance, _activeUprightVerticalText, _lengthExpressions);
}
