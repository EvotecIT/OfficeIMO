using OfficeIMO.Epub;
using OfficeIMO.Html;

internal static class ManuscriptRelationshipsFixture {
    internal static EpubPublication Create() {
        var source = HtmlConversionDocument.Parse("<html lang='en'><head><title>Manuscript relationships</title></head><body>" +
            "<section aria-labelledby='opening'><h1 id='opening'>Opening</h1>" +
            "<p aria-describedby='explanation'>This passage refers forward to supporting detail.</p>" +
            "<h1 id='detail'>Detail</h1><p id='explanation'>The supporting detail stays with its passage.</p>" +
            "<p><a href='#independent'>Continue to the independent chapter</a></p></section>" +
            "<h1 id='independent'>Independent chapter</h1><p><a href='#detail'>Return to the detail heading</a></p>" +
            "</body></html>");
        var result = EpubManuscript.ImportHtml(source, new EpubManuscriptOptions {
            Identifier = "urn:officeimo:fixture:manuscript-relationships", TypographyProfile = EpubTypographyProfile.Prose
        });
        var book = result.RequireNoLoss();
        if (book.Spine.Count != 2) throw new InvalidDataException("Related sections must share a chapter while the independent heading starts another.");
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = new IReadOnlyList<string>[] { ["textual"] },
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Text with heading navigation, a labelled section and an explicitly described passage. No images, flashing, motion or audio."
        });
        return book;
    }
}
