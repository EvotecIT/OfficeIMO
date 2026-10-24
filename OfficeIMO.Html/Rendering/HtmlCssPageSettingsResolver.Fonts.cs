using OfficeIMO.Html.Css;
using System.Threading;

namespace OfficeIMO.Html;

internal static partial class HtmlCssPageSettingsResolver {
    /// <summary>Retains authored context declarations for the shared computed cascade, including shorthand order.</summary>
    private static IReadOnlyList<HtmlGeneratedStyleDeclaration> ReadContextStyleDeclarations(string body,
        HtmlCssProcessingBudget budget, CancellationToken cancellationToken) {
        HtmlCssStyleBlock block;
        try { block = HtmlCssSyntaxParser.ParseStyleBlock(body, budget.CreateInlineSyntaxOptions(), cancellationToken); }
        catch (HtmlCssSyntaxLimitException exception) { throw budget.TranslateSyntaxLimit(exception); }
        var result = new List<HtmlGeneratedStyleDeclaration>();
        var boundedTokens = new HtmlCssTokenizationOptions { MaxInputCharacters = null, MaxTokens = null };
        int order = 0;
        foreach (HtmlCssDeclaration declaration in block.Declarations) {
            cancellationToken.ThrowIfCancellationRequested();
            string name = declaration.Name;
            if (name.StartsWith("--", StringComparison.Ordinal) || name.Equals("font", StringComparison.OrdinalIgnoreCase)
                || name.Equals("font-size", StringComparison.OrdinalIgnoreCase) || name.Equals("font-family", StringComparison.OrdinalIgnoreCase)
                || name.Equals("font-weight", StringComparison.OrdinalIgnoreCase) || name.Equals("font-style", StringComparison.OrdinalIgnoreCase)
                || name.Equals("font-stretch", StringComparison.OrdinalIgnoreCase)
                || name.Equals("line-height", StringComparison.OrdinalIgnoreCase) || name.Equals("color", StringComparison.OrdinalIgnoreCase)
                || name.Equals("text-align", StringComparison.OrdinalIgnoreCase)
                || name.Equals("vertical-align", StringComparison.OrdinalIgnoreCase)
                || IsMarginGeometryDeclaration(name) || IsMarginBackgroundDeclaration(name)) {
                string value = HtmlComputedStyleEngine.StripCssCommentsOutsideStrings(
                    HtmlCssPropertyParser.Parse(declaration, boundedTokens, cancellationToken).AuthoredValue).Trim();
                if (value.Length > 0) result.Add(new HtmlGeneratedStyleDeclaration(name, value,
                    declaration.IsImportant, 0, 0, null, order));
            }
            order++;
        }
        return result;
    }

    internal static bool IsMarginBackgroundDeclaration(string name) =>
        name.ToLowerInvariant() is "background" or "background-color" or "background-clip" or "background-image"
            or "background-position" or "background-repeat" or "background-size" or "background-origin" or "background-attachment";

    internal static bool IsMarginGeometryDeclaration(string name) =>
        name.ToLowerInvariant() is "box-sizing" or "width" or "min-width" or "max-width" or "height" or "min-height" or "max-height"
            or "margin" or "margin-top" or "margin-right" or "margin-bottom" or "margin-left" or "padding"
            or "padding-top" or "padding-right" or "padding-bottom" or "padding-left"
            or "border" or "border-width" or "border-style" or "border-color"
            or "border-top" or "border-right" or "border-bottom" or "border-left"
            or "border-top-width" or "border-right-width" or "border-bottom-width" or "border-left-width"
            or "border-top-style" or "border-right-style" or "border-bottom-style" or "border-left-style"
            or "border-top-color" or "border-right-color" or "border-bottom-color" or "border-left-color";
}
