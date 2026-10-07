using OfficeIMO.Epub;
using System.Text;

internal static class SvgPageFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("SVG fixed-layout qualification", "en", "urn:officeimo:fixture:svg-page");
        book.SetRenditionLayout(EpubRenditionLayout.PrePaginated);
        book.Creator = "OfficeIMO validation";
        book.AddResource("page", "EPUB/page.svg", "image/svg+xml", Encoding.UTF8.GetBytes("""
            <svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" version="1.1" xml:lang="en"
                 role="document" aria-labelledby="title description">
              <title id="title">An 800 by 600 SVG page</title>
              <desc id="description">Two numbered panels in document order. The first links to the second; the second links back.</desc>
              <rect x="1" y="1" width="798" height="598" fill="#fff" stroke="#172f45" stroke-width="2"/>
              <text x="40" y="65" font-family="sans-serif" font-size="30" fill="#172f45">An 800 × 600 SVG page</text>
              <g id="first" aria-label="First panel">
                <rect x="40" y="120" width="330" height="320" fill="#dcebf0"/>
                <text x="60" y="165" font-family="sans-serif" font-size="24">1. Read this first</text>
                <a xlink:href="#second"><text x="60" y="220" font-family="sans-serif" font-size="18" fill="#154e85">Continue to panel two</text></a>
              </g>
              <g id="second" aria-label="Second panel">
                <rect x="430" y="120" width="330" height="320" fill="#f3e5ce"/>
                <text x="450" y="165" font-family="sans-serif" font-size="24">2. Read this next</text>
                <a xlink:href="#first"><text x="450" y="220" font-family="sans-serif" font-size="18" fill="#154e85">Return to panel one</text></a>
              </g>
              <text x="40" y="540" font-family="sans-serif" font-size="18">Fixed canvas; reader scaling and accessibility need qualification.</text>
            </svg>
            """));
        book.AddSpineItem("page");
        book.SetNavigation(new[] { new EpubNavigationEntry("SVG page", "EPUB/page.svg") });
        book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) {
            Orientation = EpubPageOrientation.Landscape, Spread = EpubPageSpread.None, Side = EpubPageSide.Center
        });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = new[] { "visual", "textual" },
            SufficientAccessModes = new IReadOnlyList<string>[] { new[] { "visual" } },
            Features = new[] { "alternativeText", "tableOfContents" },
            Hazards = new[] { "noFlashingHazard", "noMotionSimulationHazard", "noSoundHazard" },
            Summary = "A static SVG page with selectable labels, title and description. No audio, animation or flashing. " +
                "Text does not reflow. Reader scaling, link navigation and assistive-technology reading order require independent qualification."
        });
        return book;
    }
}
