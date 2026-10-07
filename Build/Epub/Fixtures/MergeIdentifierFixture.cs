using OfficeIMO.Epub;

internal static class MergeIdentifierFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Merged identifier qualification", "en", "urn:officeimo:fixture:merge-identifiers");
        book.AddStylesheet("style", "EPUB/style.css", "html{font:18px/1.5 sans-serif;color:#202020;background:#fff}body{max-width:44rem;margin:2rem auto;padding:0 1rem}section{margin:1rem 0;padding:1rem;border:2px solid #555}#heading,#second-heading{color:#123e64}a{color:#0645ad}");
        book.AddChapter("one", "EPUB/one.xhtml", "First chapter", """
            <section aria-labelledby='heading'><h1 id='heading'>First chapter</h1><p id='description'>The original identifiers remain on this chapter.</p>
            <p><a href='two.xhtml#heading'>Go to the second heading</a></p></section>
            """, new[] { "style" });
        book.AddChapter("two", "EPUB/two.xhtml", "Second chapter", """
            <section aria-labelledby='heading'><h1 id='heading'>Second chapter</h1><p id='description'>This chapter uses the explicit replacement identifiers.</p>
            <p><a href='one.xhtml#heading'>Return to the first heading</a></p></section>
            """, new[] { "style" });
        book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions {
            SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading", ["description"] = "second-description" }
        });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Two merged text chapters with distinct heading identifiers, repaired links and section labels. No images, flashing, motion or audio."
        });
        return book;
    }
}
