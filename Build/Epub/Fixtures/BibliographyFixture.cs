using OfficeIMO.Epub;

internal static class BibliographyFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Bibliography qualification", "en", "urn:officeimo:fixture:bibliography");
        book.AddChapter("source", "EPUB/text/chapter.xhtml", "Chapter",
            "<h1>Sources</h1><p>First citation <a id='citation-one'>[1]</a>.</p><p>Repeated citation <a id='citation-two'>[1]</a>.</p>");
        book.AddChapter("references", "EPUB/back/references.xhtml", "References",
            "<section aria-labelledby='references-heading'><h1 id='references-heading'>References</h1><ol id='entries'/></section>");
        book.SetDocumentMatter("references", EpubDocumentMatter.BackMatter);
        book.AddBibliographyEntry("references", "entries", "source-one", "Example Author. <em>An illustrative source</em>. 2026.");
        book.LinkBibliographyEntry("source", "citation-one", "references", "source-one", "Return to the first citation");
        book.LinkBibliographyEntry("source", "citation-two", "references", "source-one", "Return to the second citation");
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = new[] { "textual" }, SufficientAccessModes = new IReadOnlyList<string>[] { new[] { "textual" } },
            Features = new[] { "structuralNavigation", "tableOfContents" }, Hazards = new[] { "none" },
            Summary = "A formatted bibliography entry with two citations and labelled return links. No images, flashing, motion, or audio."
        });
        return book;
    }
}
