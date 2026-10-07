using OfficeIMO.Epub;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubSvgNarrationContracts {
    private static readonly XNamespace Smil = "http://www.w3.org/ns/SMIL";
    private static readonly XNamespace Ops = "http://www.idpf.org/2007/ops";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StandaloneAndInlineSvgNarrationPreserveContentAndIdentity(bool inline) {
        var book = Book(inline, "<g id='group'><text id='first'>First</text><rect id='second' width='10' height='10'/></g>");
        byte[] content = book.GetResourceBytes("content");
        book.AddMediaOverlay("content", "overlay", "EPUB/read.smil", Overlay());
        var ids = book.GetContentXml("overlay").Descendants().Attributes("id").Select(a => a.Value).ToArray();
        book = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        book.ReplaceMediaOverlay("overlay", Overlay(1));
        Assert.Equal(content, book.GetResourceBytes("content"));
        Assert.Equal(ids, book.GetContentXml("overlay").Descendants().Attributes("id").Select(a => a.Value));
        string destination = inline ? "EPUB/new/chapter.xhtml" : "EPUB/new/page.svg";
        book.RenameResource("content", destination);
        book.RenameResource("overlay", "EPUB/new/narration.smil");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var smil = reopened.GetContentXml("overlay");
        Assert.All(smil.Descendants(Smil + "text"), e => Assert.Equal(destination,
            EpubReference.Resolve("EPUB/new/narration.smil", (string)e.Attribute("src")!).ContainerPath));
        Assert.All(smil.Descendants(Smil + "seq"), e => Assert.Equal(destination,
            EpubReference.Resolve("EPUB/new/narration.smil", (string)e.Attribute(Ops + "textref")!).ContainerPath));
        Assert.Equal(EpubPreflightStatus.Passed, reopened.Preflight().Checks.Single(c => c.Code == "media-overlays").Status);
    }

    [Theory]
    [InlineData("defs", "text")]
    [InlineData("symbol", "path")]
    [InlineData("clipPath", "rect")]
    [InlineData("mask", "rect")]
    [InlineData("pattern", "rect")]
    [InlineData("foreignObject", "text")]
    [InlineData("g", "title")]
    [InlineData("g", "desc")]
    [InlineData("g", "animate")]
    public void DefinitionMetadataAndAnimationTargetsAreRejectedAtomically(string container, string target) {
        var book = Book(false, "<" + container + "><" + target + " id='target'/></" + container + ">");
        byte[] before = book.Write().Bytes;
        var overlay = new EpubMediaOverlay { Cues = new[] { Cue("target", 0) }, AudioDurations = Durations() };
        Assert.Throws<InvalidDataException>(() => book.AddMediaOverlay("content", "overlay", "EPUB/read.smil", overlay));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData("reversed")]
    [InlineData("nested")]
    [InlineData("duplicate")]
    public void SvgCueOrderAndNonOverlapRemainRequired(string failure) {
        var book = Book(false, "<g id='group'><text id='first'>First</text><text id='second'>Second</text></g>");
        var cues = failure == "reversed" ? new[] { Cue("second", 0), Cue("first", 1) } :
            failure == "nested" ? new[] { Cue("group", 0), Cue("first", 1) } : new[] { Cue("first", 0), Cue("first", 1) };
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.AddMediaOverlay("content", "overlay", "EPUB/read.smil",
            new EpubMediaOverlay { Cues = cues, AudioDurations = Durations() }));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Book(bool inline, string body) {
        var book = EpubPublication.Create("SVG narration", "en");
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' xml:lang='en' viewBox='0 0 100 100'>" + body + "</svg>";
        if (inline) book.AddChapter("content", "EPUB/chapter.xhtml", "Chapter", svg);
        else {
            book.AddResource("content", "EPUB/page.svg", "image/svg+xml", Encoding.UTF8.GetBytes(svg));
            book.AddSpineItem("content");
            book.SetNavigation(new[] { new EpubNavigationEntry("Page", "EPUB/page.svg") });
        }
        book.AddResource("audio", "EPUB/audio.mp3", "audio/mpeg", new byte[] { 1, 2, 3 });
        return book;
    }
    private static EpubMediaOverlayCue Cue(string id, int start) => new(id, "audio", TimeSpan.FromSeconds(start), TimeSpan.FromSeconds(start + 1));
    private static Dictionary<string, TimeSpan> Durations() => new() { ["audio"] = TimeSpan.FromSeconds(10) };
    private static EpubMediaOverlay Overlay(int start = 0) => new() {
        Nodes = new[] { new EpubMediaOverlaySequence("group", new[] { Cue("first", start), Cue("second", start + 1) }) },
        AudioDurations = Durations()
    };
}
