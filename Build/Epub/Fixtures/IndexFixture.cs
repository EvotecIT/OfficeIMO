using OfficeIMO.Epub;

internal static class IndexFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Index qualification", "en", "urn:officeimo:fixture:index");
        book.AddChapter("source", "EPUB/text/chapter.xhtml", "Chapter", "<h1 id='reading-order'>Reading order</h1><p>A publication has a default reading order.</p>");
        book.AddChapter("index", "EPUB/back/index.xhtml", "Index", "<section aria-labelledby='index-heading'><h1 id='index-heading'>Index</h1><ul id='entries'/></section>");
        book.SetDocumentMatter("index", EpubDocumentMatter.BackMatter);
        book.AddIndexEntry("index", "entries", "publishing", "Publishing", Array.Empty<EpubIndexLocator>(), "publishing-entries");
        book.AddIndexEntry("index", "publishing-entries", "reading", "reading order", new[] {
            new EpubIndexLocator { ManifestId = "source", FragmentId = "reading-order", Label = "Reading order" },
            new EpubIndexLocator { ManifestId = "source", Label = "Whole chapter" }
        });
        book.AddIndexEntry("index", "entries", "books", "Books — see", new[] {
            new EpubIndexLocator { ManifestId = "index", FragmentId = "publishing", Label = "Publishing" }
        });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = new[] { "textual" }, SufficientAccessModes = new IReadOnlyList<string>[] { new[] { "textual" } },
            Features = new[] { "structuralNavigation", "tableOfContents", "index" }, Hazards = new[] { "none" },
            Summary = "A nested index with labelled links to a section, a chapter and another index entry. No images, flashing, motion, or audio."
        });
        return book;
    }
}
