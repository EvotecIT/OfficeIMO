using OfficeIMO.Epub;
using OfficeIMO.Bibliography;
using System.Xml.Linq;

internal static class BibliographyFixture {
    internal static EpubPublication Create() {
        var document = BibliographyDocument.Parse("""
            [{"id":"source-one","type":"book","title":"Research & Publishing","author":[{"family":"Example","given":"Ada"}],"issued":{"date-parts":[[2026]]}},
             {"id":"source-two","type":"book","title":"Accessible Books","author":[{"family":"Writer","given":"Alex"}],"issued":{"date-parts":[[2025]]}}]
            """, BibliographyFormat.CslJson).Document;
        var style = CslStyle.Parse("""
            <style xmlns="http://purl.org/net/xbiblio/csl" version="1.0" class="in-text">
              <citation><layout prefix="[" suffix="]" delimiter=", "><text variable="citation-number"/></layout></citation>
              <bibliography><sort><key variable="title"/></sort><layout suffix=".">
                <names variable="author" suffix=". "><name/></names>
                <text variable="title" font-style="italic" suffix=". "/>
                <date variable="issued"><date-part name="year"/></date>
              </layout></bibliography>
            </style>
            """);
        var first = new CslCitation("citation-one"); first.Items.Add(new CslCitationItem("source-one"));
        var second = new CslCitation("citation-two"); second.Items.Add(new CslCitationItem("source-one"));
        var third = new CslCitation("citation-three"); third.Items.Add(new CslCitationItem("source-two"));
        var rendered = new CslProcessor(document, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html })
            .Render(new[] { first, second, third });
        if (rendered.Bibliography.Count != 2 || rendered.Bibliography[0].Key != "source-two")
            throw new InvalidDataException("The fixture requires title-sorted CSL bibliography output.");
        var book = EpubPublication.Create("Bibliography qualification", "en", "urn:officeimo:fixture:bibliography");
        book.AddChapter("source", "EPUB/text/chapter.xhtml", "Chapter",
            "<h1>Sources</h1>" + string.Concat(rendered.Citations.Select(citation =>
                "<p>Citation <a id=\"" + citation.Key + "\">" + citation.Content + "</a>.</p>")));
        book.AddChapter("references", "EPUB/back/references.xhtml", "References",
            "<section aria-labelledby='references-heading'><h1 id='references-heading'>References</h1><ol id='entries'/></section>");
        book.SetDocumentMatter("references", EpubDocumentMatter.BackMatter);
        foreach (var entry in rendered.Bibliography) book.AddBibliographyEntry("references", "entries", entry.Key, entry.Content);
        book.LinkBibliographyEntry("source", "citation-one", "references", "source-one", "Return to the first citation");
        book.LinkBibliographyEntry("source", "citation-two", "references", "source-one", "Return to the second citation");
        book.LinkBibliographyEntry("source", "citation-three", "references", "source-two", "Return to the third citation");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        XNamespace html = "http://www.w3.org/1999/xhtml";
        var entries = reopened.GetContentXml("references").Descendants(html + "li").ToArray();
        if (entries.Length != 2 || entries[0].Attribute("id")?.Value != "source-two" ||
            !entries[1].Value.Contains("Research & Publishing", StringComparison.Ordinal) ||
            !entries[1].Descendants().Any(element => element.Name == html + "i" || element.Name == html + "em" ||
                ((string?)element.Attribute("style"))?.Contains("italic", StringComparison.Ordinal) == true))
            throw new InvalidDataException("CSL sorting, escaped text, or italic formatting did not survive EPUB reopening.");
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = new[] { "textual" }, SufficientAccessModes = new IReadOnlyList<string>[] { new[] { "textual" } },
            Features = new[] { "structuralNavigation", "tableOfContents" }, Hazards = new[] { "none" },
            Summary = "Two CSL-formatted bibliography entries with three citations and labelled return links. No images, flashing, motion, or audio."
        });
        return book;
    }
}
