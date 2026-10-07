using OfficeIMO.Epub;
using System.Text;
using System.Xml.Linq;

internal static class CssPreservationFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("CSS preservation qualification", "en", "urn:officeimo:fixture:css-preservation");
        book.AddResource("texture", "EPUB/texture.svg", "image/svg+xml", Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xml:lang='en' width='2' height='2'><rect width='2' height='2' fill='#f4f8fc'/></svg>"));
        book.AddStylesheet("style", "EPUB/style.css", """
            /* Keep this comment and the compound selector intact. */
            html{font:18px/1.5 sans-serif;color:#202020;background:white}
            body{max-width:42rem;margin:2rem auto;padding:0 1rem}
            #first/**/.lead{color:#123e64;border-left:6px solid currentColor;padding:1rem;background:url('texture.svg')}
            a{color:#0645ad}
            """);
        book.AddChapter("one", "EPUB/one.xhtml", "First chapter", """
            <section aria-labelledby='first'><h1 id='first' class='lead'>First chapter</h1>
            <p>The blue heading keeps its color, border and background after the image moves.</p>
            <p><a href='parts/two.xhtml#second'>Continue to the second chapter</a></p></section>
            """, ["style"]);
        book.AddChapter("two", "EPUB/parts/two.xhtml", "Second chapter", """
            <section aria-labelledby='second'><h1 id='second' class='lead'>Second chapter</h1>
            <p>The green heading keeps its color, border and background after the chapter merge.</p>
            <p><a href='../one.xhtml#first'>Return to the first chapter</a></p></section>
            """, ["style"]);
        XNamespace html = "http://www.w3.org/1999/xhtml";
        var second = book.GetContentXml("two");
        second.Root!.Element(html + "head")!.Add(new XElement(html + "style",
            "#second/**/.lead{color:#185b3a;border-left:6px solid currentColor;padding:1rem;background:url('../texture.svg')} /* retained */"));
        book.SetContentXml("two", second);
        book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions { StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles });
        book.RenameResource("texture", "EPUB/art/texture.svg");
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Two text chapters with headings and repaired navigation links. Background images are decorative; no flashing, motion or audio."
        });
        return book;
    }
}
