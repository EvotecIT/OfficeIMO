using OfficeIMO.Epub;
using System.Text;

internal static class SvgNarrationFixture {
    internal static EpubPublication Create(bool inline) {
        var book = EpubPublication.Create("SVG narration qualification", "en", "urn:officeimo:fixture:svg-narration:" + inline);
        book.Creator = "OfficeIMO validation";
        const string svg = """
            <svg xmlns="http://www.w3.org/2000/svg" version="1.1" xml:lang="en" width="1200" height="300" viewBox="0 0 1200 300"
                 role="document" aria-labelledby="title description">
              <title id="title">Two narrated SVG sentences</title>
              <desc id="description">Two sentences in document order with recorded narration. Reader playback is not yet qualified.</desc>
              <style type="text/css">text { font-family: sans-serif; font-size: 24px; fill: #172f45; } .narration-active { fill: #8c2600; }</style>
              <rect width="1200" height="300" fill="#fff"/>
              <g id="sentences">
                <text id="first" x="40" y="100">First, the reader hears this sentence.</text>
                <text id="second" x="40" y="200">Next, the narration continues with the second sentence.</text>
              </g>
            </svg>
            """;
        if (inline) book.AddChapter("content", "EPUB/page.xhtml", "Two narrated SVG sentences", svg);
        else {
            book.AddResource("content", "EPUB/page.svg", "image/svg+xml", Encoding.UTF8.GetBytes(svg));
            book.AddSpineItem("content");
            book.SetNavigation(new[] { new EpubNavigationEntry("Two narrated SVG sentences", "EPUB/page.svg") });
            book.SetRenditionLayout(EpubRenditionLayout.PrePaginated);
            book.SetFixedLayoutPage("content", new EpubFixedLayoutPage(1200, 300) { Spread = EpubPageSpread.None, Side = EpubPageSide.Center });
        }
        using var stream = typeof(SvgNarrationFixture).Assembly.GetManifestResourceStream("OfficeIMO.Epub.Fixtures.narration.mp3")!;
        using var bytes = new MemoryStream(); stream.CopyTo(bytes);
        book.AddResource("audio", "EPUB/audio/narration.mp3", "audio/mpeg", bytes.ToArray());
        EpubMediaOverlay Overlay(double boundary) => new() {
            AudioDurations = new Dictionary<string, TimeSpan> { ["audio"] = TimeSpan.FromTicks(71179590) },
            Nodes = new[] { new EpubMediaOverlaySequence("sentences", new[] {
                new EpubMediaOverlayCue("first", "audio", TimeSpan.Zero, TimeSpan.FromSeconds(boundary)),
                new EpubMediaOverlayCue("second", "audio", TimeSpan.FromSeconds(boundary), TimeSpan.FromTicks(71179590))
            }) }
        };
        book.AddMediaOverlay("content", "narration", "EPUB/overlays/narration.smil", Overlay(2.8));
        book.ReplaceMediaOverlay("narration", Overlay(2.85));
        book.SetMetadataProperty("media:active-class", "narration-active");
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = new[] { "textual", "visual", "auditory" },
            SufficientAccessModes = new IReadOnlyList<string>[] { new[] { "textual" } },
            Features = new[] { "alternativeText", "tableOfContents", "synchronizedAudioText" },
            Hazards = new[] { "unknown" },
            Summary = "Two selectable SVG sentences have recorded narration. The page title and description are not narrated. Text does not reflow. Native scaling, highlighting, playback, assistive technology and hazard review remain unqualified."
        });
        return book;
    }
}
