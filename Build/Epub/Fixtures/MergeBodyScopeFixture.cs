using OfficeIMO.Epub;
using System.Xml.Linq;

internal static class MergeBodyScopeFixture {
    internal static EpubPublication Create(bool merged) {
        XNamespace html = "http://www.w3.org/1999/xhtml";
        var book = EpubPublication.Create("Body scope merge qualification", "en", "urn:officeimo:fixture:merge-body-scopes");
        book.AddChapter("one", "EPUB/one.xhtml", "First chapter", "<h1 id='first-heading'>First chapter</h1><p>The first chapter keeps its navy text and pale blue background. This longer sentence should wrap cleanly in a narrow reading window.</p><p><a href='two.xhtml#second-body'>Continue to the second chapter</a></p>");
        book.AddChapter("two", "EPUB/two.xhtml", "Second chapter", "<h1 id='second-heading'>Second chapter</h1><p>The second chapter keeps its maroon text and pale pink background. Each chapter remains an independently addressable content scope.</p><p><a href='one.xhtml#first-body'>Return to the first chapter</a></p>");
        foreach (var item in new[] { ("one", "first-body", "first", "color:#102a43;background-color:#eaf2f8;padding:12px;border:2px solid #102a43"),
            ("two", "second-body", "second", "color:#702020;background-color:#fff1f1;padding:12px;border:2px solid #702020") }) {
            var document = book.GetContentXml(item.Item1);var body = document.Root!.Element(html + "body")!;
            body.SetAttributeValue("id", item.Item2);body.SetAttributeValue("class", item.Item3);body.SetAttributeValue("style", item.Item4);
            body.SetAttributeValue("data-chapter", item.Item1);
            book.SetContentXml(item.Item1, document);
        }
        if (merged) book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions { PreserveBodyScopes = true });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Two text chapters with reciprocal navigation links. No images, flashing, motion or audio."
        });
        return book;
    }
}
