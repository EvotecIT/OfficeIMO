using OfficeIMO.Epub;

internal static class NotesAndPagesFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Notes and print-page qualification", "en", "urn:officeimo:fixture:notes-and-pages");
        book.AddStylesheet("style", "EPUB/style.css", EpubTypography.CreateStylesheet(EpubTypographyProfile.Prose));
        book.AddChapter("preface", "EPUB/front/preface.xhtml", "Preface",
            "<h1>Preface</h1><p><span id='page-iv'/>This synthetic edition exercises front matter, notes and source-page navigation.</p>", ["style"]);
        book.AddChapter("chapter", "EPUB/text/chapter.xhtml", "Chapter",
            "<h1>Chapter</h1><p><span id='page-one'/>This passage has a same-document footnote <a id='footnote-ref'>1</a>.</p>" +
            "<p>This passage has a cross-document endnote <a id='endnote-ref'>2</a>.</p>" +
            "<section aria-labelledby='footnotes-heading'><h2 id='footnotes-heading'>Footnotes</h2><div id='footnotes'/></section>", ["style"]);
        book.AddChapter("notes", "EPUB/back/notes.xhtml", "Endnotes",
            "<section aria-labelledby='endnotes-heading'><h1 id='endnotes-heading'>Endnotes</h1>" +
            "<p><span id='page-two'/>Supporting material for the chapter.</p><ol id='notes-list'/></section>", ["style"]);
        book.SetDocumentMatter("preface", EpubDocumentMatter.FrontMatter);
        book.SetDocumentMatter("chapter", EpubDocumentMatter.BodyMatter);
        book.SetDocumentMatter("notes", EpubDocumentMatter.BackMatter);
        book.AddPrintPageMarker("preface", "page-iv", "iv", "Reference pages");
        book.AddPrintPageMarker("chapter", "page-one", "1", "Reference pages");
        book.AddPrintPageMarker("notes", "page-two", "2", "Reference pages");
        book.SetMetadataProperty("pageBreakSource", "urn:officeimo:fixture:synthetic-print-edition");
        book.AddNote(new EpubNoteOptions {
            SourceManifestId = "chapter", ReferenceId = "footnote-ref", NotesManifestId = "chapter", ContainerId = "footnotes",
            NoteId = "footnote-1", Kind = EpubNoteKind.Footnote, BodyXhtml = "<p>Additional detail in the same document.</p>",
            BacklinkText = "Return to footnote reference 1"
        });
        book.AddNote(new EpubNoteOptions {
            SourceManifestId = "chapter", ReferenceId = "endnote-ref", NotesManifestId = "notes", ContainerId = "notes-list",
            NoteId = "endnote-2", Kind = EpubNoteKind.Endnote, BodyXhtml = "<p>Supporting detail in the back matter.</p>",
            BacklinkText = "Return to endnote reference 2"
        });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = new IReadOnlyList<string>[] { ["textual"] },
            Features = ["structuralNavigation", "tableOfContents", "pageNavigation"], Hazards = ["none"],
            Summary = "Text with labelled footnote and endnote links, return links, and three synthetic print-page markers. No images, flashing, motion or audio."
        });
        return book;
    }
}
