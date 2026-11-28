#region QuickStart
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;

HtmlDocument source = HtmlDocumentEngine.Default.ParseDocument(
    "<h1>Service review</h1><table><tbody><tr id='items'><th>Item</th></tr></tbody></table>");
HtmlElement row = source.QuerySelector("#items")!;
HtmlDocumentFragment cells = HtmlDocumentEngine.Default.ParseFragment(
    "<td>Certificates</td><td>Renewed</td>", row);

HtmlDocument edited = source.Edit(document => {
    document.QuerySelector("#items")!.AppendChild(document.ImportNode(cells));
});

HtmlConversionDocument report = HtmlConversionDocument.FromDocument(edited);
File.WriteAllText("Service-review.html", report.ExportSourceHtml());
#endregion
