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
                || name.Equals("all", StringComparison.OrdinalIgnoreCase)
                || name.Equals("font-size", StringComparison.OrdinalIgnoreCase) || name.Equals("font-family", StringComparison.OrdinalIgnoreCase)
                || name.Equals("font-weight", StringComparison.OrdinalIgnoreCase) || name.Equals("font-style", StringComparison.OrdinalIgnoreCase)
                || name.Equals("line-height", StringComparison.OrdinalIgnoreCase) || name.Equals("color", StringComparison.OrdinalIgnoreCase)
                || name.Equals("text-align", StringComparison.OrdinalIgnoreCase)) {
                string value = HtmlComputedStyleEngine.StripCssCommentsOutsideStrings(
                    HtmlCssPropertyParser.Parse(declaration, boundedTokens, cancellationToken).AuthoredValue).Trim();
                if (value.Length > 0) result.Add(new HtmlGeneratedStyleDeclaration(name, value,
                    declaration.IsImportant, 0, 0, null, order));
            }
            order++;
        }
        return result;
    }
}
