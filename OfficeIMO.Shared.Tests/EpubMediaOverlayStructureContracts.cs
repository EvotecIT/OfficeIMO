using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMediaOverlayStructureContracts {
    private static readonly XNamespace Smil = "http://www.w3.org/ns/SMIL";
    private static readonly XNamespace Ops = "http://www.idpf.org/2007/ops";

    [Fact]
    public void NestedSequencesRetainSemanticsIdentityAndReferencesThroughReplaceAndRename() {
        var book = Book();
        book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Tree());
        book = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var original = book.GetContentXml("overlay");
        Assert.Equal(new[] { "aside", "list" }, original.Descendants(Smil + "seq").Skip(1).Select(e => (string?)e.Attribute(Ops + "type")));
        Assert.Equal("footnote", (string?)original.Descendants(Smil + "par").Last().Attribute(Ops + "type"));
        var ids = original.Descendants().Attributes("id").Select(a => a.Value).ToArray();
        book.ReplaceMediaOverlay("overlay", Tree(2));
        Assert.Equal(ids, book.GetContentXml("overlay").Descendants().Attributes("id").Select(a => a.Value));
        Assert.Equal("2s", (string?)book.GetContentXml("overlay").Descendants(Smil + "audio").First().Attribute("clipBegin"));
        book.RenameResource("chapter", "EPUB/text/new.xhtml");
        book.RenameResource("overlay", "EPUB/overlays/new.smil");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.All(reopened.GetContentXml("overlay").Descendants(Smil + "seq"), seq =>
            Assert.Equal("EPUB/text/new.xhtml", EpubReference.Resolve("EPUB/overlays/new.smil", (string)seq.Attribute(Ops + "textref")!).ContainerPath));
        Assert.Equal(EpubPreflightStatus.Passed, reopened.Preflight().Checks.Single(c => c.Code == "media-overlays").Status);
        Assert.All(reopened.GetPackageXml().Descendants().Where(e => (string?)e.Attribute("property") == "media:duration"), e => Assert.Equal("3s", e.Value));
    }

    [Fact]
    public void RemovingReferencedSequenceIsAtomicWhileSurvivingSequenceRemainsReplaceable() {
        var book = Book(); book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Tree());
        book.AddStylesheet("style", "EPUB/style.css", "p { background: url('read.smil#seq0'); }");
        byte[] before = book.Write().Bytes;
        var flat = new EpubMediaOverlay { Cues = new[] { Cue("first"), Cue("second"), Cue("note") }, AudioDurations = Durations() };
        Assert.Throws<InvalidOperationException>(() => book.ReplaceMediaOverlay("overlay", flat));
        Assert.Equal(before, book.Write().Bytes);
        book.ReplaceMediaOverlay("overlay", Tree(1));
        Assert.Contains(book.GetContentXml("overlay").Descendants().Attributes("id"), a => a.Value == "seq0");
    }

    [Theory]
    [InlineData("outside")]
    [InlineData("order")]
    [InlineData("duplicate")]
    [InlineData("semantic")]
    [InlineData("mixed")]
    public void InvalidTreesFailWithoutMutatingPublication(string failure) {
        var book = Book(); byte[] before = book.Write().Bytes;
        var overlay = Tree();
        if (failure == "outside") overlay.Nodes = new[] { new EpubMediaOverlaySequence("items", new[] { Cue("note") }) };
        if (failure == "order") overlay.Nodes = new EpubMediaOverlayNode[] { Cue("second"), Cue("first") };
        if (failure == "duplicate") overlay.Nodes = new EpubMediaOverlayNode[] { Cue("first"), Cue("first") };
        if (failure == "semantic") overlay.Nodes[0].Semantic = (EpubMediaOverlaySemantic)999;
        if (failure == "mixed") overlay.Cues = new[] { Cue("first") };
        if (failure == "semantic") Assert.Throws<ArgumentOutOfRangeException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", overlay));
        else if (failure == "mixed") Assert.Throws<ArgumentException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", overlay));
        else Assert.Throws<InvalidDataException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", overlay));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void DepthAndTotalNodeBoundsRejectReachableLargeTrees() {
        var book = EpubPublication.Create("Nested", "en");
        string body = "<p id='leaf'>Text</p>";
        EpubMediaOverlayNode node = Cue("leaf");
        for (int i = 0; i < 33; i++) {
            body = "<section id='s" + i + "'>" + body + "</section>";
            node = new EpubMediaOverlaySequence("s" + i, new[] { node });
        }
        book.AddChapter("chapter", "EPUB/chapter.xhtml", "Chapter", body);
        book.AddResource("audio", "EPUB/audio.mp3", "audio/mpeg", new byte[] { 1 });
        byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil",
            new EpubMediaOverlay { Nodes = new[] { node }, AudioDurations = Durations() }));
        Assert.Equal(before, book.Write().Bytes);
        var large = Enumerable.Range(0, 10001).Select(i => (EpubMediaOverlayNode)Cue("leaf")).ToArray();
        Assert.Throws<ArgumentException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil",
            new EpubMediaOverlay { Nodes = large, AudioDurations = Durations() }));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Nested", "en");
        book.AddChapter("chapter", "EPUB/chapter.xhtml", "Chapter", "<aside id='box'><ul id='items'><li id='first'>First</li><li id='second'>Second</li></ul><p id='note'>Note</p></aside>");
        book.AddResource("audio", "EPUB/audio.mp3", "audio/mpeg", new byte[] { 1, 2, 3 });
        return book;
    }
    private static EpubMediaOverlayCue Cue(string id, int start = 0) => new(id, "audio", TimeSpan.FromSeconds(start), TimeSpan.FromSeconds(start + 1));
    private static Dictionary<string, TimeSpan> Durations() => new() { ["audio"] = TimeSpan.FromSeconds(10) };
    private static EpubMediaOverlay Tree(int start = 0) => new() {
        AudioDurations = Durations(),
        Nodes = new[] { new EpubMediaOverlaySequence("box", new EpubMediaOverlayNode[] {
            new EpubMediaOverlaySequence("items", new[] { Cue("first", start), Cue("second", start + 1) }) { Semantic = EpubMediaOverlaySemantic.List },
            new EpubMediaOverlayCue("note", "audio", TimeSpan.FromSeconds(start + 2), TimeSpan.FromSeconds(start + 3)) { Semantic = EpubMediaOverlaySemantic.Footnote }
        }) { Semantic = EpubMediaOverlaySemantic.Aside } }
    };
}
