using OfficeIMO.Epub;

internal static class MergeMatterFixture {
    internal static EpubPublication Create(bool merged) {
        var book = EpubPublication.Create("Document matter merge qualification", "en", "urn:officeimo:fixture:merge-matter");
        book.AddChapter("one", "EPUB/one.xhtml", "Preface", "<h1 id='preface'>Preface</h1><p>This introduction belongs to the front matter.</p><p><a href='two.xhtml#chapter'>Begin the chapter</a></p>");
        book.AddChapter("two", "EPUB/two.xhtml", "Main chapter", "<h1 id='chapter'>Main chapter</h1><p>This chapter belongs to the body matter.</p><p><a href='one.xhtml#preface'>Return to the preface</a></p>");
        book.SetDocumentMatter("one", EpubDocumentMatter.FrontMatter);
        book.SetDocumentMatter("two", EpubDocumentMatter.BodyMatter);
        if (merged) book.MergeChapters("one", "two", "chapter-start", new EpubChapterMergeOptions { PreserveDocumentMatter = true });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "A preface and main chapter with reciprocal navigation links. No images, flashing, motion or audio."
        });
        return book;
    }
}
