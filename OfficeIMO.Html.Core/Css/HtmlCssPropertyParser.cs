using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Html.Css;

/// <summary>Parses selected property values over the owned CSS tokenizer.</summary>
public static class HtmlCssPropertyParser {
    private static readonly HashSet<string> DisplayKeywords = Set(
        "block", "inline", "inline-block", "none", "flex", "inline-flex", "grid", "inline-grid",
        "table", "table-caption", "table-column-group", "table-column", "table-header-group",
        "table-row-group", "table-footer-group", "table-row", "table-cell", "list-item", "contents", "flow-root", "-webkit-box");
    private static readonly HashSet<string> VisibilityKeywords = Set("visible", "hidden", "collapse");
    private static readonly HashSet<string> AutoKeyword = Set("auto");
    private static readonly HashSet<string> NoneKeyword = Set("none");
    private static readonly HashSet<string> NonNegativeLengthProperties = Set(
        "width", "height", "min-width", "min-height", "max-width", "max-height",
        "padding-top", "padding-right", "padding-bottom", "padding-left");
    private static readonly HashSet<string> SystemColors = Set(
        "accentcolor", "accentcolortext", "activetext", "buttonborder", "buttonface", "buttontext",
        "canvas", "canvastext", "field", "fieldtext", "graytext", "highlight", "highlighttext",
        "linktext", "mark", "marktext", "selecteditem", "selecteditemtext", "visitedtext");
    private static readonly HashSet<string> NamedColors = Set(
        "aliceblue", "antiquewhite", "aqua", "aquamarine", "azure", "beige", "bisque", "black", "blanchedalmond",
        "blue", "blueviolet", "brown", "burlywood", "cadetblue", "chartreuse", "chocolate", "coral", "cornflowerblue",
        "cornsilk", "crimson", "cyan", "darkblue", "darkcyan", "darkgoldenrod", "darkgray", "darkgreen", "darkgrey",
        "darkkhaki", "darkmagenta", "darkolivegreen", "darkorange", "darkorchid", "darkred", "darksalmon", "darkseagreen",
        "darkslateblue", "darkslategray", "darkslategrey", "darkturquoise", "darkviolet", "deeppink", "deepskyblue",
        "dimgray", "dimgrey", "dodgerblue", "firebrick", "floralwhite", "forestgreen", "fuchsia", "gainsboro", "ghostwhite",
        "gold", "goldenrod", "gray", "green", "greenyellow", "grey", "honeydew", "hotpink", "indianred", "indigo",
        "ivory", "khaki", "lavender", "lavenderblush", "lawngreen", "lemonchiffon", "lightblue", "lightcoral", "lightcyan",
        "lightgoldenrodyellow", "lightgray", "lightgreen", "lightgrey", "lightpink", "lightsalmon", "lightseagreen",
        "lightskyblue", "lightslategray", "lightslategrey", "lightsteelblue", "lightyellow", "lime", "limegreen", "linen",
        "magenta", "maroon", "mediumaquamarine", "mediumblue", "mediumorchid", "mediumpurple", "mediumseagreen",
        "mediumslateblue", "mediumspringgreen", "mediumturquoise", "mediumvioletred", "midnightblue", "mintcream", "mistyrose",
        "moccasin", "navajowhite", "navy", "oldlace", "olive", "olivedrab", "orange", "orangered", "orchid",
        "palegoldenrod", "palegreen", "paleturquoise", "palevioletred", "papayawhip", "peachpuff", "peru", "pink", "plum",
        "powderblue", "purple", "rebeccapurple", "red", "rosybrown", "royalblue", "saddlebrown", "salmon", "sandybrown",
        "seagreen", "seashell", "sienna", "silver", "skyblue", "slateblue", "slategray", "slategrey", "snow", "springgreen",
        "steelblue", "tan", "teal", "thistle", "tomato", "transparent", "turquoise", "violet", "wheat", "white",
        "whitesmoke", "yellow", "yellowgreen");

    /// <summary>Parses a standalone property value with explicit lexical budgets.</summary>
    public static HtmlCssPropertyParseResult Parse(
        string propertyName,
        string value,
        HtmlCssTokenizationOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (propertyName == null) throw new ArgumentNullException(nameof(propertyName));
        if (value == null) throw new ArgumentNullException(nameof(value));
        ValidateInputLength(value, options);
        string authoredValue = value.Trim();
        IReadOnlyList<HtmlCssToken> tokens = HtmlCssTokenizer.Tokenize(authoredValue, options, cancellationToken);
        return ParseTokens(propertyName.Trim(), authoredValue, tokens, isImportant: false, cancellationToken);
    }

    /// <summary>Parses a source-preserving declaration and removes its trailing !important annotation.</summary>
    public static HtmlCssPropertyParseResult Parse(HtmlCssDeclaration declaration, CancellationToken cancellationToken = default) =>
        Parse(declaration, options: null, cancellationToken);

    /// <summary>Parses a source-preserving declaration with explicit lexical budgets.</summary>
    public static HtmlCssPropertyParseResult Parse(
        HtmlCssDeclaration declaration,
        HtmlCssTokenizationOptions? options,
        CancellationToken cancellationToken = default) {
        if (declaration == null) throw new ArgumentNullException(nameof(declaration));
        cancellationToken.ThrowIfCancellationRequested();
        string value = declaration.ValueText;
        ValidateInputLength(value, options);
        if (declaration.IsImportant) value = RemoveImportantSuffix(value, options, cancellationToken);
        string authoredValue = value.Trim();
        IReadOnlyList<HtmlCssToken> tokens = HtmlCssTokenizer.Tokenize(authoredValue, options, cancellationToken);
        return ParseTokens(declaration.Name, authoredValue, tokens, declaration.IsImportant, cancellationToken);
    }

    private static HtmlCssPropertyParseResult ParseTokens(
        string propertyName,
        string authoredValue,
        IReadOnlyList<HtmlCssToken> tokens,
        bool isImportant,
        CancellationToken cancellationToken) {
        if (!HtmlCssPropertyCatalog.TryGet(propertyName, out HtmlCssPropertyDefinition? definition))
            return Result(propertyName, authoredValue, HtmlCssPropertyParseStatus.UnknownProperty, null, null, isImportant);
        List<HtmlCssToken> significant = tokens.Where(token => !IsTrivia(token.Kind) && token.Kind != HtmlCssTokenKind.EndOfFile).ToList();
        if (significant.Count == 0 || HasInvalidStructure(significant, cancellationToken))
            return Result(propertyName, authoredValue, HtmlCssPropertyParseStatus.InvalidSyntax, definition, null, isImportant);
        if (TryCssWideKeyword(authoredValue, significant, out HtmlCssPropertyValue? wide))
            return Result(propertyName, authoredValue, HtmlCssPropertyParseStatus.Parsed, definition, wide, isImportant);
        if (ContainsVar(significant)) {
            if (!HasValidVarFunctions(significant, cancellationToken))
                return Result(propertyName, authoredValue, HtmlCssPropertyParseStatus.InvalidSyntax, definition, null, isImportant);
            var deferred = new HtmlCssPropertyValue(HtmlCssPropertyValueKind.DeferredFunction, authoredValue, authoredValue);
            return Result(propertyName, authoredValue, HtmlCssPropertyParseStatus.Deferred, definition, deferred, isImportant);
        }

        HtmlCssPropertyValue? parsed = definition!.Name switch {
            "display" => ParseKeyword(authoredValue, significant, DisplayKeywords),
            "visibility" => ParseKeyword(authoredValue, significant, VisibilityKeywords),
            "opacity" => ParseOpacity(authoredValue, significant, cancellationToken),
            "color" => ParseColor(authoredValue, significant, cancellationToken),
            "width" or "height" or "min-width" or "min-height" =>
                ParseLengthProperty(definition.Name, authoredValue, significant, AutoKeyword, cancellationToken),
            "max-width" or "max-height" =>
                ParseLengthProperty(definition.Name, authoredValue, significant, NoneKeyword, cancellationToken),
            "margin-top" or "margin-right" or "margin-bottom" or "margin-left" =>
                ParseLengthProperty(definition.Name, authoredValue, significant, AutoKeyword, cancellationToken),
            "padding-top" or "padding-right" or "padding-bottom" or "padding-left" =>
                ParseLengthProperty(definition.Name, authoredValue, significant, null, cancellationToken),
            _ => null
        };
        return Result(propertyName, authoredValue,
            parsed == null ? HtmlCssPropertyParseStatus.UnsupportedValue : HtmlCssPropertyParseStatus.Parsed,
            definition, parsed, isImportant);
    }

    private static HtmlCssPropertyValue? ParseKeyword(string text, IReadOnlyList<HtmlCssToken> tokens, ISet<string> keywords) {
        if (tokens.Count != 1 || tokens[0].Kind != HtmlCssTokenKind.Identifier) return null;
        string canonical = (tokens[0].Value ?? string.Empty).ToLowerInvariant();
        return keywords.Contains(canonical)
            ? new HtmlCssPropertyValue(HtmlCssPropertyValueKind.Keyword, text, canonical)
            : null;
    }

    private static HtmlCssPropertyValue? ParseOpacity(string text, IReadOnlyList<HtmlCssToken> tokens, CancellationToken cancellationToken) {
        if (!HtmlCssTypedValueParsers.TryParseNumeric(text, tokens, cancellationToken, out HtmlCssNumericValue? numeric)) return null;
        double number = numeric!.Value;
        bool percentage = numeric.Type == HtmlCssNumericType.Percentage;
        return new HtmlCssPropertyValue(
            numeric.IsCalculated ? HtmlCssPropertyValueKind.Calculation
                : percentage ? HtmlCssPropertyValueKind.Percentage : HtmlCssPropertyValueKind.Number,
            text,
            number.ToString("R", CultureInfo.InvariantCulture) + (percentage ? "%" : string.Empty),
            number,
            numericValue: numeric);
    }

    private static HtmlCssPropertyValue? ParseLengthProperty(
        string propertyName,
        string text,
        IReadOnlyList<HtmlCssToken> tokens,
        ISet<string>? keywords,
        CancellationToken cancellationToken) {
        HtmlCssPropertyValue? keyword = keywords == null ? null : ParseKeyword(text, tokens, keywords);
        if (keyword != null) return keyword;
        HtmlCssMathParseResult parsed;
        try {
            parsed = HtmlCssMathParser.ParseLengthPercentage(text, tokens, cancellationToken);
        } catch (HtmlCssMathLimitException) {
            return null;
        }
        if (!parsed.IsParsed) return null;
        HtmlCssMathExpression expression = parsed.Expression!;
        if (NonNegativeLengthProperties.Contains(propertyName)
            && expression.Kind == HtmlCssMathExpressionKind.Literal
            && expression.Value < 0D) return null;
        HtmlCssPropertyValueKind kind = expression.IsCalculated
            ? HtmlCssPropertyValueKind.Calculation
            : expression.Type == HtmlCssNumericType.Percentage
                ? HtmlCssPropertyValueKind.Percentage
                : HtmlCssPropertyValueKind.Length;
        return new HtmlCssPropertyValue(kind, text, expression.CanonicalText,
            expression.Kind == HtmlCssMathExpressionKind.Literal ? expression.Value : null,
            mathExpression: expression);
    }

    private static HtmlCssPropertyValue? ParseColor(string text, IReadOnlyList<HtmlCssToken> tokens, CancellationToken cancellationToken) {
        if (tokens.Count > 1) {
            if (!HtmlCssTypedValueParsers.TryParseColorFunction(text, tokens, cancellationToken,
                    out HtmlCssColorFunctionValue? function)) return null;
            int open = text.IndexOf('(');
            string functionCanonical = open < 0 ? text : text.Substring(0, open).Trim().ToLowerInvariant() + text.Substring(open);
            return new HtmlCssPropertyValue(HtmlCssPropertyValueKind.ColorFunction, text, functionCanonical, colorFunction: function);
        }
        if (tokens.Count != 1) return null;
        HtmlCssToken token = tokens[0];
        if (token.Kind == HtmlCssTokenKind.Hash) {
            string digits = token.Value ?? string.Empty;
            if ((digits.Length == 3 || digits.Length == 4 || digits.Length == 6 || digits.Length == 8)
                && digits.All(IsHexDigit))
                return new HtmlCssPropertyValue(HtmlCssPropertyValueKind.HexColor, text, "#" + digits.ToLowerInvariant());
            return null;
        }
        if (token.Kind != HtmlCssTokenKind.Identifier) return null;
        string canonical = (token.Value ?? string.Empty).ToLowerInvariant();
        if (canonical == "currentcolor") return new HtmlCssPropertyValue(HtmlCssPropertyValueKind.CurrentColor, text, canonical);
        if (SystemColors.Contains(canonical)) return new HtmlCssPropertyValue(HtmlCssPropertyValueKind.SystemColor, text, canonical);
        return NamedColors.Contains(canonical)
            ? new HtmlCssPropertyValue(HtmlCssPropertyValueKind.NamedColor, text, canonical)
            : null;
    }

    private static bool TryCssWideKeyword(string text, IReadOnlyList<HtmlCssToken> tokens, out HtmlCssPropertyValue? value) {
        value = null;
        if (tokens.Count != 1 || tokens[0].Kind != HtmlCssTokenKind.Identifier) return false;
        string canonical = (tokens[0].Value ?? string.Empty).ToLowerInvariant();
        HtmlCssWideKeyword keyword;
        switch (canonical) {
            case "initial": keyword = HtmlCssWideKeyword.Initial; break;
            case "inherit": keyword = HtmlCssWideKeyword.Inherit; break;
            case "unset": keyword = HtmlCssWideKeyword.Unset; break;
            case "revert": keyword = HtmlCssWideKeyword.Revert; break;
            case "revert-layer": keyword = HtmlCssWideKeyword.RevertLayer; break;
            default: return false;
        }
        value = new HtmlCssPropertyValue(HtmlCssPropertyValueKind.CssWideKeyword, text, canonical, cssWideKeyword: keyword);
        return true;
    }

    private static bool ContainsVar(IReadOnlyList<HtmlCssToken> tokens) =>
        tokens.Any(token => token.Kind == HtmlCssTokenKind.Function && string.Equals(token.Value, "var", StringComparison.OrdinalIgnoreCase));

    private static bool HasValidVarFunctions(IReadOnlyList<HtmlCssToken> tokens, CancellationToken cancellationToken) {
        for (int index = 0; index < tokens.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (tokens[index].Kind != HtmlCssTokenKind.Function
                || !string.Equals(tokens[index].Value, "var", StringComparison.OrdinalIgnoreCase)) continue;
            int depth = 1;
            int cursor = index + 1;
            int first = -1;
            int afterName = -1;
            for (; cursor < tokens.Count && depth > 0; cursor++) {
                HtmlCssToken token = tokens[cursor];
                if (token.Kind == HtmlCssTokenKind.Function || token.Kind == HtmlCssTokenKind.OpenParenthesis
                    || token.Kind == HtmlCssTokenKind.OpenBracket || token.Kind == HtmlCssTokenKind.OpenBrace) {
                    depth++;
                    continue;
                }
                if (token.Kind == HtmlCssTokenKind.CloseParenthesis || token.Kind == HtmlCssTokenKind.CloseBracket
                    || token.Kind == HtmlCssTokenKind.CloseBrace) {
                    depth--;
                    continue;
                }
                if (depth != 1 || IsTrivia(token.Kind)) continue;
                if (first < 0) first = cursor;
                else if (afterName < 0) afterName = cursor;
            }
            if (depth != 0 || first < 0 || tokens[first].Kind != HtmlCssTokenKind.Identifier
                || tokens[first].Value?.StartsWith("--", StringComparison.Ordinal) != true) return false;
            if (afterName >= 0 && tokens[afterName].Kind != HtmlCssTokenKind.Comma) return false;
        }
        return true;
    }

    private static bool HasInvalidStructure(IReadOnlyList<HtmlCssToken> tokens, CancellationToken cancellationToken) {
        var closers = new Stack<HtmlCssTokenKind>();
        foreach (HtmlCssToken token in tokens) {
            cancellationToken.ThrowIfCancellationRequested();
            if (token.Kind == HtmlCssTokenKind.BadString || token.Kind == HtmlCssTokenKind.BadUrl) return true;
            switch (token.Kind) {
                case HtmlCssTokenKind.Function:
                case HtmlCssTokenKind.OpenParenthesis: closers.Push(HtmlCssTokenKind.CloseParenthesis); break;
                case HtmlCssTokenKind.OpenBracket: closers.Push(HtmlCssTokenKind.CloseBracket); break;
                case HtmlCssTokenKind.OpenBrace: closers.Push(HtmlCssTokenKind.CloseBrace); break;
                case HtmlCssTokenKind.CloseParenthesis:
                case HtmlCssTokenKind.CloseBracket:
                case HtmlCssTokenKind.CloseBrace:
                    if (closers.Count == 0 || closers.Pop() != token.Kind) return true;
                    break;
            }
        }
        return closers.Count != 0;
    }

    private static string RemoveImportantSuffix(
        string value,
        HtmlCssTokenizationOptions? options,
        CancellationToken cancellationToken) {
        IReadOnlyList<HtmlCssToken> tokens = HtmlCssTokenizer.Tokenize(value, options, cancellationToken);
        int last = tokens.Count - 2;
        while (last >= 0 && IsTrivia(tokens[last].Kind)) last--;
        if (last < 0 || tokens[last].Kind != HtmlCssTokenKind.Identifier
            || !string.Equals(tokens[last].Value, "important", StringComparison.OrdinalIgnoreCase)) return value;
        int bang = last - 1;
        while (bang >= 0 && IsTrivia(tokens[bang].Kind)) bang--;
        if (bang < 0 || tokens[bang].Kind != HtmlCssTokenKind.Delimiter || tokens[bang].Value != "!") return value;
        return value.Substring(0, tokens[bang].Offset).TrimEnd();
    }

    private static void ValidateInputLength(string value, HtmlCssTokenizationOptions? options) {
        HtmlCssTokenizationOptions effective = (options ?? new HtmlCssTokenizationOptions()).Clone();
        effective.Validate();
        if (effective.MaxInputCharacters.HasValue && value.Length > effective.MaxInputCharacters.Value)
            throw new HtmlCssTokenizationLimitException(nameof(effective.MaxInputCharacters), value.Length, effective.MaxInputCharacters.Value);
    }

    private static HtmlCssPropertyParseResult Result(string propertyName, string value, HtmlCssPropertyParseStatus status,
        HtmlCssPropertyDefinition? definition, HtmlCssPropertyValue? parsed, bool important) =>
        new HtmlCssPropertyParseResult(propertyName, value, status, definition, parsed, important);

    private static bool IsTrivia(HtmlCssTokenKind kind) => kind == HtmlCssTokenKind.Whitespace || kind == HtmlCssTokenKind.Comment;
    private static bool IsHexDigit(char value) => value >= '0' && value <= '9' || value >= 'a' && value <= 'f' || value >= 'A' && value <= 'F';
    private static HashSet<string> Set(params string[] values) => new HashSet<string>(values, StringComparer.OrdinalIgnoreCase);
}
