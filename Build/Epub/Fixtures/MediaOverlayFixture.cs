using OfficeIMO.Epub;
using System.Xml.Linq;

internal static class MediaOverlayFixture {
    internal static EpubPublication Nested(bool revised = false) {
        var book = Create();
        XNamespace html = "http://www.w3.org/1999/xhtml";
        var content = book.GetContentXml("chapter");
        var paragraphs = content.Descendants(html + "p").ToArray();
        var list = new XElement(html + "ol", new XAttribute("id", "steps"));
        foreach (var paragraph in paragraphs) list.Add(new XElement(html + "li", paragraph.Attributes(), paragraph.Nodes()));
        paragraphs[0].AddBeforeSelf(new XElement(html + "section", new XAttribute("id", "narrated"), list));
        foreach (var paragraph in paragraphs) paragraph.Remove();
        book.SetContentXml("chapter", content);
        EpubMediaOverlay Overlay(double boundary) => new() {
            AudioDurations = new Dictionary<string, TimeSpan> { ["audio"] = TimeSpan.FromTicks(71179590) },
            Nodes = new[] { new EpubMediaOverlaySequence("narrated", new[] {
                new EpubMediaOverlaySequence("steps", new[] {
                    new EpubMediaOverlayCue("first", "audio", TimeSpan.Zero, TimeSpan.FromSeconds(boundary)) { Semantic = EpubMediaOverlaySemantic.ListItem },
                    new EpubMediaOverlayCue("second", "audio", TimeSpan.FromSeconds(boundary), TimeSpan.FromTicks(71179590)) { Semantic = EpubMediaOverlaySemantic.ListItem }
                }) { Semantic = EpubMediaOverlaySemantic.List }
            }) }
        };
        book.ReplaceMediaOverlay("narration", Overlay(2.8));
        if (revised) book.ReplaceMediaOverlay("narration", Overlay(2.85));
        book.SetMetadataProperty("schema:accessibilitySummary", "All content is available as text. Two list items have recorded narration in nested SMIL sequences. The heading is not narrated. Reader playback, escaping and hazard review are not yet qualified.");
        return book;
    }

    internal static EpubPublication Revised() {
        var book = Create();
        // Move the cue boundary within the independently measured inter-sentence silence.
        book.ReplaceMediaOverlay("narration", new EpubMediaOverlay {
            AudioDurations = new Dictionary<string, TimeSpan> { ["audio"] = TimeSpan.FromTicks(71179590) },
            Cues = new[] {
                new EpubMediaOverlayCue("first", "audio", TimeSpan.Zero, TimeSpan.FromSeconds(2.85)),
                new EpubMediaOverlayCue("second", "audio", TimeSpan.FromSeconds(2.85), TimeSpan.FromTicks(71179590))
            }
        });
        return book;
    }

    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Read-aloud qualification", "en", "urn:officeimo:fixture:read-aloud");
        book.Creator = "OfficeIMO validation";
        book.AddStylesheet("style", "EPUB/style.css", "body { font: 1.25em serif; line-height: 1.6; } .narration-active { background: #ffe090; color: #111; }");
        book.AddChapter("chapter", "EPUB/text/chapter.xhtml", "Two narrated sentences",
            "<h1>Two narrated sentences</h1><p id='first'>First, the reader hears this sentence.</p><p id='second'>Next, the narration continues with the second sentence.</p>", new[] { "style" });
        using var stream = typeof(MediaOverlayFixture).Assembly.GetManifestResourceStream("OfficeIMO.Epub.Fixtures.narration.mp3")!;
        using var bytes = new MemoryStream(); stream.CopyTo(bytes);
        book.AddResource("audio", "EPUB/audio/narration.mp3", "audio/mpeg", bytes.ToArray());
        book.AddMediaOverlay("chapter", "narration", "EPUB/overlays/narration.smil", new EpubMediaOverlay {
            AudioDurations = new Dictionary<string, TimeSpan> { ["audio"] = TimeSpan.FromTicks(71179590) },
            Cues = new[] {
                new EpubMediaOverlayCue("first", "audio", TimeSpan.Zero, TimeSpan.FromSeconds(2.8)),
                new EpubMediaOverlayCue("second", "audio", TimeSpan.FromSeconds(2.8), TimeSpan.FromTicks(71179590))
            }
        });
        book.SetMetadataProperty("media:active-class", "narration-active");
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = new[] { "textual", "auditory" },
            SufficientAccessModes = new IReadOnlyList<string>[] { new[] { "textual" } },
            Features = new[] { "structuralNavigation", "tableOfContents", "synchronizedAudioText" },
            Hazards = new[] { "unknown" },
            Summary = "All content is available as text. Two paragraphs have recorded narration and SMIL text associations. " +
                "The heading is not narrated. Reader playback, synchronization and hazard review are not yet qualified."
        });
        return book;
    }
}
