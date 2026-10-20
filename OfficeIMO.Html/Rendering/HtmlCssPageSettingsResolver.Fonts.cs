using AngleSharp.Css.Parser;
using OfficeIMO.Html.Css;
using System.Threading;

namespace OfficeIMO.Html;

internal static partial class HtmlCssPageSettingsResolver {
    /// <summary>Replays the font family before the retained parser can collapse or discard declarations.</summary>
    private static IReadOnlyDictionary<string, HtmlCssPageDeclaration> ReadMarginFontDeclarations(string body,
        HtmlCssProcessingBudget budget, CancellationToken cancellationToken) {
        var result = new Dictionary<string, HtmlCssPageDeclaration>(StringComparer.OrdinalIgnoreCase);
        int order = 0;
        HtmlCssStyleBlock block;
        try { block = HtmlCssSyntaxParser.ParseStyleBlock(body, budget.CreateInlineSyntaxOptions(), cancellationToken); }
        catch (HtmlCssSyntaxLimitException exception) { throw budget.TranslateSyntaxLimit(exception); }
        var boundedTokens = new HtmlCssTokenizationOptions { MaxInputCharacters = null, MaxTokens = null };
        foreach (HtmlCssDeclaration declaration in block.Declarations) {
            cancellationToken.ThrowIfCancellationRequested();
            string name = declaration.Name.ToLowerInvariant();
            if (name != "font" && name != "font-size" && name != "line-height" && name != "font-family"
                && name != "font-weight" && name != "font-style") { order++; continue; }
            string value = HtmlComputedStyleEngine.StripCssCommentsOutsideStrings(
                HtmlCssPropertyParser.Parse(declaration, boundedTokens, cancellationToken).AuthoredValue).Trim();
            if (name == "font") {
                if (HtmlComputedStyleEngine.TryExpandFontShorthand(value, out IReadOnlyList<KeyValuePair<string, string>> expanded)) {
                    foreach (KeyValuePair<string, string> item in expanded) Consider(item.Key, item.Value, declaration.IsImportant);
                }
            } else if (name == "font-size" || name == "line-height") {
                if (IsRevertLayer(value) || IsValidMarginFontLength(name, value)) Consider(name, value, declaration.IsImportant);
            } else if (name == "font-family" || name == "font-weight" || name == "font-style") {
                string parsed = new CssParser().ParseDeclaration(name + ":" + value)?.GetPropertyValue(name) ?? string.Empty;
                if (IsRevertLayer(value)) parsed = value;
                if (parsed.Length > 0) Consider(name, parsed, declaration.IsImportant);
            }
            order++;
        }
        return result;

        void Consider(string name, string value, bool important) {
            if (!result.TryGetValue(name, out HtmlCssPageDeclaration earlier) || important || !earlier.IsImportant)
                result[name] = new HtmlCssPageDeclaration(value, important, order);
        }
    }

    private static bool IsValidMarginFontLength(string name, string value) {
        // Keep ordinary keywords/provider grammar, and accept owned lengths only
        // through the same validator used by computed font longhands.
        return HtmlComputedStyleEngine.IsSupportedDeclarationValue(name, value);
    }

    private static HtmlCssPageDeclaration ReadMarginFontDeclaration(
        IReadOnlyDictionary<string, HtmlCssPageDeclaration> declarations, string name, int order) =>
        declarations.TryGetValue(name, out HtmlCssPageDeclaration declaration)
            ? declaration : new HtmlCssPageDeclaration(string.Empty, false, order);
}
