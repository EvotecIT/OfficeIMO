using OfficeIMO.Epub;
using System.Text;
using System.Xml.Linq;

internal static class MergeResourceSelectorsFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Merge resource selector qualification", "en", "urn:officeimo:fixture:merge-resource-selectors");
        book.AddStylesheet("shared", "EPUB/shared.css", """
            @import "styles/links.css";
            html { font:18px/1.5 sans-serif; color:#202020; background:white }
            body { max-width:42rem; margin:2rem auto; padding:0 1rem }
            section { margin-bottom:2rem; padding:1rem; border:1px solid #777 }
            img { width:80px; height:80px }
            """);
        book.AddStylesheet("links", "EPUB/styles/links.css", """
            a[href='#heading'] { color:#123e64; font-weight:bold; border-bottom:3px solid currentColor }
            """);
        book.AddResource("cover", "EPUB/cover.svg", "image/svg+xml", Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xml:lang='en' width='80' height='80' viewBox='0 0 80 80'><rect width='80' height='80' fill='#e8eef5'/><circle cx='40' cy='40' r='24' fill='#123e64'/></svg>"));
        book.AddChapter("one", "EPUB/one.xhtml", "First chapter",
            "<section><h1 id='heading'>First chapter</h1><p><a href='#heading'>First local link</a></p><p><a class='forward' href='parts/two.xhtml#heading'>Continue to second chapter</a></p></section>", ["shared"]);
        book.AddChapter("two", "EPUB/parts/two.xhtml", "Second chapter",
            "<section><h1 id='heading'>Second chapter</h1><p><a href='#heading'>Second local link</a></p>" +
            "<blockquote cite='two.xhtml#heading'><p>This quotation retains its green styling.</p></blockquote>" +
            "<img src='../cover.svg' alt='Blue circle on a pale background'/><p><a href='../three.xhtml#heading'>Continue to third chapter</a></p></section>", ["shared"]);
        book.AddChapter("three", "EPUB/three.xhtml", "Third chapter",
            "<section><h1 id='heading'>Third chapter</h1><p><a href='#heading'>Third local link</a></p><p><a href='parts/two.xhtml#heading'>Return to second chapter</a></p></section>", ["shared"]);
        XNamespace html = "http://www.w3.org/1999/xhtml";
        var first = book.GetContentXml("one");
        first.Root!.Element(html + "head")!.Add(new XElement(html + "style",
            "a.forward[href='parts/two.xhtml#heading'] { color:#704000; border:2px dotted currentColor }"));
        book.SetContentXml("one", first);
        var second = book.GetContentXml("two");
        second.Root!.Element(html + "head")!.Add(new XElement(html + "style", """
            a[href='#heading'], blockquote[cite='two.xhtml#heading'] { color:#185b3a }
            img[src='../cover.svg'] { border:4px solid #185b3a }
            """));
        book.SetContentXml("two", second);
        book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions {
            StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteChapterSelectors = true,
            SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading" }
        });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual", "visual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents", "alternativeText"], Hazards = ["none"],
            Summary = "Text chapters, descriptive image alternative text and repaired links. No flashing, motion or audio."
        });
        return book;
    }
}
