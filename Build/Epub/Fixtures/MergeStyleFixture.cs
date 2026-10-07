using OfficeIMO.Epub;
using System.Text;
using System.Xml.Linq;

internal static class MergeStyleFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Merged stylesheet qualification", "en", "urn:officeimo:fixture:merge-styles");
        book.AddResource("common", "EPUB/styles/common.css", "text/css", Encoding.UTF8.GetBytes(
            "html{font:18px/1.5 sans-serif;color:#202020;background:#fff}body{max-width:44rem;margin:2rem auto;padding:0 1rem}p{color:#222}a{color:#0645ad}section{padding:1rem;border:2px solid #555;margin:1rem 0}"));
        book.AddChapter("one", "EPUB/text/one.xhtml", "First chapter", """
            <section><h1 id='first'>First chapter</h1><p id='first-text'>Both chapters use the final purple paragraph color.</p>
            <p><a href='../parts/two.xhtml#second'>Go to the second chapter</a></p></section>
            """);
        book.AddChapter("two", "EPUB/parts/two.xhtml", "Second chapter", """
            <section><h1 id='second'>Second chapter</h1><p id='second-text'>The appended rule participates in one shared cascade.</p>
            <p><a href='../text/one.xhtml#first'>Return to the first chapter</a></p></section>
            """);
        XNamespace html = "http://www.w3.org/1999/xhtml";
        foreach (string id in new[] { "one", "two" }) {
            var xml = book.GetContentXml(id);
            xml.Root!.Element(html + "head")!.Add(new XElement(html + "link", new XAttribute("rel", "stylesheet"), new XAttribute("href", "../styles/common.css")),
                new XElement(html + "style", id == "one" ? "p{color:#003366}section{border-radius:8px}" : "p{color:#663366}section{border-style:dashed}"));
            book.SetContentXml(id, xml);
        }
        book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions { StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Two merged text chapters with repaired navigation and explicitly combined styles. No images, flashing, motion or audio."
        });
        return book;
    }
}
