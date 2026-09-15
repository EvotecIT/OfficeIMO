using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;

var options = new HtmlParseOptions {
    MaxInputCharacters = 16 * 1024,
    MaxNodes = 256,
    MaxDepth = 32
};
HtmlDocument source = AngleSharpHtmlParser.Instance.ParseDocument(
    "<table><tbody><tr id='items'><th>Item</th></tr></tbody></table>", options);
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
