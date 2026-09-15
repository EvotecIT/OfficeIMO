using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Css;
using OfficeIMO.Html.Providers;

var options = new HtmlParseOptions {
    MaxInputCharacters = 16 * 1024,
    MaxNodes = 256,
    MaxDepth = 32
};
HtmlDocument source = AngleSharpHtmlParser.Instance.ParseDocument(
    "<table><tbody><tr id='items'><th>Item</th></tr></tbody></table><svg><a id='asset'></a></svg>", options);
HtmlElement row = source.QuerySelector("#items")
    ?? throw new InvalidOperationException("The packed document query lost its table row.");
HtmlDocumentFragment cells = AngleSharpHtmlParser.Instance.ParseFragment(
    "<td>Quarterly report</td><td>Approved</td>", row, options);
HtmlDocument edited = source.Edit(document => {
    HtmlElement target = document.QuerySelector("#items")!;
    target.AppendChild(document.ImportNode(cells));
});

if (source.QuerySelectorAll("td").Count != 0 ||
    edited.QuerySelectorAll("td").Count != 2 ||
    edited.OuterHtml.IndexOf("Quarterly report", StringComparison.Ordinal) < 0) {
    throw new InvalidOperationException("The packed document parse, fragment, edit or serialization contract failed.");
}

const string css = "@future report; .status { color: green; future-property: fn(one[two]); }";
HtmlCssStyleSheet styleSheet = HtmlCssSyntaxParser.ParseStyleSheet(css);
HtmlCssStyleBlock inlineStyle = HtmlCssSyntaxParser.ParseStyleBlock("color:green;color:future");
HtmlCssPropertyParseResult parsedOpacity = HtmlCssPropertyParser.Parse("opacity", "calc(20% + 55%)");
HtmlCssPropertyParseResult parsedColor = HtmlCssPropertyParser.Parse("color", "hsl(210 50 40 / 75%)");
HtmlCssSelectorParseResult parsedSelector = HtmlCssSelectorParser.Parse("table > tbody tr#items");
HtmlCssStyleSheet namespaceSheet = HtmlCssSyntaxParser.ParseStyleSheet(
    "@namespace svg url('http://www.w3.org/2000/svg');svg|a:first-child{}");
var selectorOptions = new HtmlCssSelectorOptions {
    Namespaces = HtmlCssNamespaceContext.FromStyleSheet(namespaceSheet)
};
HtmlCssSelectorListParseResult parsedSelectorList = HtmlCssSelectorParser.ParseList(
    "#missing, svg|a:first-child:is(#asset):not(.missing)", selectorOptions);
HtmlCssMathParseResult parsedWidth = HtmlCssMathParser.ParseLengthPercentage("calc(24px + 25%)");
HtmlCssLengthResolutionResult resolvedWidth = HtmlCssMathResolver.ResolveLength(parsedWidth.Expression!,
    new HtmlCssLengthResolutionContext { PercentageReference = 200D });
if (styleSheet.Rules.Count != 2 || styleSheet.ToCss() != css ||
    inlineStyle.Declarations.Count != 2 || inlineStyle.Declarations[1].Name != "color" ||
    parsedOpacity.Status != HtmlCssPropertyParseStatus.Parsed || parsedOpacity.Value?.NumericValue?.Value != 75D ||
    parsedOpacity.Value.NumericValue.IsCalculated != true || parsedColor.Value?.ColorFunction?.Kind != HtmlCssColorFunctionKind.Hsl ||
    parsedColor.Value.ColorFunction.Alpha.Value != 75D || parsedSelector.Selector?.Matches(edited.QuerySelector("#items")!) != true ||
    parsedSelectorList.SelectorList?.Matches(edited.QuerySelector("#asset")!) != true ||
    parsedWidth.Expression?.Type != HtmlCssNumericType.LengthPercentage || resolvedWidth.Value != 74D) {
    throw new InvalidOperationException("The packed CSS syntax or property-grammar contract failed.");
}

string[] coreReferences = typeof(HtmlDocument).Assembly.GetReferencedAssemblies()
    .Select(reference => reference.Name ?? string.Empty)
    .ToArray();
string[] providerReferences = typeof(AngleSharpHtmlParser).Assembly.GetReferencedAssemblies()
    .Select(reference => reference.Name ?? string.Empty)
    .ToArray();
if (coreReferences.Any(reference => reference.StartsWith("AngleSharp", StringComparison.Ordinal)
        || reference == "OfficeIMO.Core"
        || reference == "OfficeIMO.Html") ||
    providerReferences.Contains("OfficeIMO.Html", StringComparer.Ordinal)
        || providerReferences.Contains("OfficeIMO.Core", StringComparer.Ordinal)) {
    throw new InvalidOperationException("The packed standalone document graph pulled in the conversion or graphics layer.");
}

Console.WriteLine("OfficeIMO HTML document-only packed API smoke passed on " +
    System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription + ".");
