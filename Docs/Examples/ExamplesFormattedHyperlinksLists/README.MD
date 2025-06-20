## Hyperlink lists with formatting reuse

This example demonstrates how to build two bullet lists of hyperlinks where the second list reuses the formatting from the first one.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var first = document.AddList(WordListStyle.Bulleted);
    var google = first.AddItem("").AddHyperLink("Google", new Uri("https://google.com"), addStyle: true);
    google.Bold = true;
    var bing = first.AddItem("").AddHyperLink("Bing", new Uri("https://bing.com"), addStyle: true);
    bing.Italic = true;
    var yahoo = first.AddItem("").AddHyperLink("Yahoo", new Uri("https://yahoo.com"), addStyle: true);
    yahoo.Color = Color.Purple;

    document.AddParagraph("Some paragraph separating the lists.");
    document.AddParagraph("Another paragraph.");

    var second = document.AddList(WordListStyle.Bulleted);
    second.AddItem("").AddHyperLink("DuckDuckGo", new Uri("https://duckduckgo.com")).Hyperlink
        .CopyFormattingFrom(google.Hyperlink);
    second.AddItem("").AddHyperLink("Startpage", new Uri("https://startpage.com")).Hyperlink
        .CopyFormattingFrom(bing.Hyperlink);
    second.AddItem("").AddHyperLink("GitHub", new Uri("https://github.com")).Hyperlink
        .CopyFormattingFrom(yahoo.Hyperlink);

    document.Save();
}
```
