using OfficeIMO.Epub;

internal static class GlossaryFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Glossary qualification", "en", "urn:officeimo:fixture:glossary");
        book.AddStylesheet("style", "EPUB/style.css", EpubTypography.CreateStylesheet(EpubTypographyProfile.Prose));
        book.AddChapter("source", "EPUB/text/chapter.xhtml", "Chapter",
            "<h1>Publishing terms</h1><p>A <a id='ref-one'>reflowable book</a> adapts to reader settings.</p>" +
            "<p>This second passage also mentions a <a id='ref-two'>reflowable book</a>.</p>", new[] { "style" });
        book.AddChapter("glossary", "EPUB/back/glossary.xhtml", "Glossary",
            "<section aria-labelledby='glossary-heading'><h1 id='glossary-heading'>Glossary</h1><dl id='terms'/></section>", new[] { "style" });
        book.SetDocumentMatter("glossary", EpubDocumentMatter.BackMatter);
        book.AddGlossaryEntry("glossary", "terms", "reflowable", "Reflowable book",
            "<p>A publication whose text can adapt to the available reading area and the reader’s font settings.</p>");
        book.AddGlossaryEntry("glossary", "terms", "spine", "Spine", "<p>The publication’s default reading order.</p>");
        book.LinkGlossaryTerm("source", "ref-one", "glossary", "reflowable", "Return to the first passage");
        book.LinkGlossaryTerm("source", "ref-two", "glossary", "reflowable", "Return to the second passage");
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = new[] { "textual" }, SufficientAccessModes = new IReadOnlyList<string>[] { new[] { "textual" } },
            Features = new[] { "structuralNavigation", "tableOfContents" }, Hazards = new[] { "none" },
            Summary = "Two glossary terms with links from the chapter and labelled return links. No images, flashing, motion, or audio."
        });
        return book;
    }
}
