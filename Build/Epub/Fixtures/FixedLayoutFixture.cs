using OfficeIMO.Epub;

internal static class FixedLayoutFixture {
    internal static EpubPublication Create(bool rightToLeft, bool packageLayout = true) {
        var book = EpubPublication.Create(packageLayout ? "Fixed-layout package qualification" : "Fixed-layout item override qualification", "en", "urn:officeimo:fixture:fixed-layout:" + (rightToLeft ? "rtl" : "ltr") + (packageLayout ? ":package" : ":item"));
        if (packageLayout) book.SetRenditionLayout(EpubRenditionLayout.PrePaginated);
        book.PageProgressionDirection = rightToLeft ? "rtl" : "ltr";
        book.Creator = "OfficeIMO validation";
        book.AddChapter("landscape", "EPUB/landscape.xhtml", "Landscape canvas",
            "<header id='page-header'><h1 id='title'>An 800 × 600 page</h1></header>" +
            "<section id='first-region' aria-labelledby='first-heading' style='background:#dcebf0;padding:12px'>" +
            "<h2 id='first-heading'>1. Read this first</h2><p>The DOM keeps this section first, regardless of page placement.</p><p><a href='#second-heading'>Continue to section two</a></p></section>" +
            "<section id='second-region' aria-labelledby='second-heading' style='background:#f3e5ce;padding:12px'>" +
            "<h2 id='second-heading'>2. Read this next</h2><p>Text remains selectable and semantic.</p><p><a href='#first-heading'>Return to section one</a></p></section>" +
            "<footer id='page-footer'><p>Canvas boundary: 800 × 600 CSS pixels.</p></footer>");
        book.AddChapter("portrait", "EPUB/portrait.xhtml", "Portrait canvas",
            "<main id='portrait-main'><h1>A 600 × 800 page</h1><p>This page requests centered, single-page presentation.</p><p><a href='landscape.xhtml#title'>Return to the landscape page</a></p></main>");
        book.SetFixedLayoutPage("landscape", new EpubFixedLayoutPage(800, 600) {
            Orientation = EpubPageOrientation.Landscape, Spread = EpubPageSpread.Both,
            Side = rightToLeft ? EpubPageSide.Left : EpubPageSide.Right,
            Regions = new[] {
                new EpubFixedLayoutRegion("page-header", 40, 32, 720, 90),
                new EpubFixedLayoutRegion("first-region", 40, 150, 330, 300),
                new EpubFixedLayoutRegion("second-region", 430, 150, 330, 300),
                new EpubFixedLayoutRegion("page-footer", 40, 520, 720, 60)
            }
        });
        book.SetFixedLayoutPage("portrait", new EpubFixedLayoutPage(600, 800) {
            Orientation = EpubPageOrientation.Portrait, Spread = EpubPageSpread.None, Side = EpubPageSide.Center,
            Regions = new[] { new EpubFixedLayoutRegion("portrait-main", 40, 40, 520, 720) }
        });
        return book;
    }

    internal static EpubPublication EscapedIdentifier() {
        var book = EpubPublication.Create("Positioned identifier qualification", "en", "urn:officeimo:fixture:fixed-layout:escaped-id");
        book.SetRenditionLayout(EpubRenditionLayout.PrePaginated);
        book.AddChapter("page", "EPUB/page.xhtml", "Identifier-safe placement",
            "<section id='ordinary'><h1>Identifier-safe placement</h1><p>This box is positioned by a quoted CSS attribute selector.</p></section>");
        var document = book.GetContentXml("page");
        const string actualId = "panel:\"\\😀</style>";
        document.Descendants(System.Xml.Linq.XName.Get("section", "http://www.w3.org/1999/xhtml")).Single().SetAttributeValue("id", actualId);
        book.SetContentXml("page", document);
        book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) {
            Spread = EpubPageSpread.None,
            Regions = new[] { new EpubFixedLayoutRegion(actualId, 40.125m, 50.5m, 320.25m, 240.75m) }
        });
        return book;
    }

}
