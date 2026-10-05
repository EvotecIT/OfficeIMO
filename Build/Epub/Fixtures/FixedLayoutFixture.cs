using OfficeIMO.Epub;

internal static class FixedLayoutFixture {
    internal static EpubPublication Create(bool rightToLeft, bool packageLayout = true) {
        var book = EpubPublication.Create(packageLayout ? "Fixed-layout package qualification" : "Fixed-layout item override qualification", "en", "urn:officeimo:fixture:fixed-layout:" + (rightToLeft ? "rtl" : "ltr") + (packageLayout ? ":package" : ":item"));
        if (packageLayout) book.SetRenditionLayout(EpubRenditionLayout.PrePaginated);
        book.PageProgressionDirection = rightToLeft ? "rtl" : "ltr";
        book.Creator = "OfficeIMO validation";
        book.AddChapter("landscape", "EPUB/landscape.xhtml", "Landscape canvas",
            "<header style='position:absolute;left:40px;top:32px;width:720px'><h1 id='title'>An 800 × 600 page</h1></header>" +
            "<section aria-labelledby='first-heading' style='position:absolute;left:40px;top:150px;width:330px;height:300px;background:#dcebf0;padding:12px;box-sizing:border-box'>" +
            "<h2 id='first-heading'>1. Read this first</h2><p>The DOM keeps this section first, regardless of page placement.</p><p><a href='#second-heading'>Continue to section two</a></p></section>" +
            "<section aria-labelledby='second-heading' style='position:absolute;left:430px;top:150px;width:330px;height:300px;background:#f3e5ce;padding:12px;box-sizing:border-box'>" +
            "<h2 id='second-heading'>2. Read this next</h2><p>Text remains selectable and semantic.</p><p><a href='#first-heading'>Return to section one</a></p></section>" +
            "<footer style='position:absolute;left:40px;top:520px;width:720px'><p>Canvas boundary: 800 × 600 CSS pixels.</p></footer>");
        book.AddChapter("portrait", "EPUB/portrait.xhtml", "Portrait canvas",
            "<main style='position:absolute;left:40px;top:40px;width:520px'><h1>A 600 × 800 page</h1><p>This page requests centered, single-page presentation.</p><p><a href='landscape.xhtml#title'>Return to the landscape page</a></p></main>");
        book.SetFixedLayoutPage("landscape", new EpubFixedLayoutPage(800, 600) {
            Orientation = EpubPageOrientation.Landscape, Spread = EpubPageSpread.Both,
            Side = rightToLeft ? EpubPageSide.Left : EpubPageSide.Right
        });
        book.SetFixedLayoutPage("portrait", new EpubFixedLayoutPage(600, 800) {
            Orientation = EpubPageOrientation.Portrait, Spread = EpubPageSpread.None, Side = EpubPageSide.Center
        });
        return book;
    }
}
