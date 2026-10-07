using OfficeIMO.Epub;

internal static class SplitFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Chapter split qualification", "en", "urn:officeimo:fixture:split");
        book.AddChapter("source", "EPUB/text/chapter.xhtml", "One chapter", """
            <section id='chapter'><h1 id='first'>First section</h1>
            <p>A chapter can become two reading positions. <a href='#second'>Continue to the second section</a>.</p>
            <h1 id='second'>Second section</h1><p>This section moves to a new resource. <a href='#first'>Return to the first section</a>.</p></section>
            """);
        book.AddChapter("references", "EPUB/back/references.xhtml", "References", "<h1>References</h1><p><a href='../text/chapter.xhtml#second'>Second section</a></p>");
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = new[] { "textual" }, SufficientAccessModes = new IReadOnlyList<string>[] { new[] { "textual" } },
            Features = new[] { "structuralNavigation", "tableOfContents" }, Hazards = new[] { "none" },
            Summary = "Two linked sections split into consecutive chapters, with a repaired reference from another chapter. No images, flashing, motion, or audio."
        });
        book.SplitChapter("source", "second", "second", "EPUB/parts/second.xhtml", "Second section");
        return book;
    }
}
