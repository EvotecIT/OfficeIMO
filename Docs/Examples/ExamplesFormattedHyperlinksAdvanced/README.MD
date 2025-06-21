## Advanced formatted hyperlinks

This example demonstrates how to duplicate hyperlinks, insert new links around an existing one and reuse the formatting.
It also shows that the helper methods work in headers and footers.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddHeadersAndFooters();
    var paragraph = document.AddParagraph("Visit ");
    paragraph.AddHyperLink("Google", new Uri("https://google.com"), addStyle: true);
    var baseLink = paragraph.Hyperlink;
    baseLink.Bold = true;

    baseLink.InsertFormattedHyperlinkBefore("Bing", new Uri("https://bing.com"));
    var duplicate = WordHyperLink.DuplicateHyperlink(baseLink);
    duplicate.Text = "Google Copy";

    var yahoo = baseLink.InsertFormattedHyperlinkAfter("Yahoo", new Uri("https://yahoo.com"));
    yahoo.CopyFormattingFrom(baseLink);

    var headerPara = document.Header.Default.AddParagraph("Search with ");
    var duck = headerPara.AddHyperLink("DuckDuckGo", new Uri("https://duckduckgo.com"), addStyle: true);
    duck.Hyperlink.InsertFormattedHyperlinkAfter("Startpage", new Uri("https://startpage.com"));

    var footerPara = document.Footer.Default.AddParagraph("Code on ");
    var gitHub = footerPara.AddHyperLink("GitHub", new Uri("https://github.com"), addStyle: true);
    gitHub.Hyperlink.InsertFormattedHyperlinkBefore("GitLab", new Uri("https://gitlab.com"));

    document.Save();
}
```
