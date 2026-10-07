using System.Threading;
using OfficeIMO.Epub;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMediaOverlayReplacementContracts {
    private static readonly XNamespace Smil = "http://www.w3.org/ns/SMIL";
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";

    [Fact]
    public void SameDurationReplacementPersistsAndKeepsResourceIdentityAndAudio() {
        var book = Book();
        byte[] original = book.Write().Bytes;
        book = EpubPublication.Load(new MemoryStream(original));
        var item = book.Manifest.Single(i => i.Id == "overlay");
        book.ReplaceMediaOverlay("overlay", Overlay(Cue("first", 4), Cue("second", 6)));
        Assert.Equal("EPUB/read.smil", item.Reference.ContainerPath);
        Assert.Equal("overlay", book.Manifest.Single(i => i.Id == "chapter").MediaOverlayId);
        byte[] revised = book.Write().Bytes;
        Assert.NotEqual(original, revised);
        var reopened = EpubPublication.Load(new MemoryStream(revised));
        Assert.Equal(new[] { "4s", "6s" }, reopened.GetContentXml("overlay").Descendants(Smil + "audio").Select(e => (string?)e.Attribute("clipBegin")));
        Assert.Equal(new[] { "cue0", "cue1" }, CueIds(reopened));
        Assert.Equal(ReadEntry(original, "EPUB/audio.mp3"), ReadEntry(revised, "EPUB/audio.mp3"));
        Assert.All(Durations(reopened), e => Assert.Equal("2s", e.Value));
    }

    [Fact]
    public void SurvivingTargetsKeepIdsAndNewTargetsNeverReuseRemovedCueIds() {
        var book = Book();
        book.ReplaceMediaOverlay("overlay", Overlay(Cue("second"), Cue("third")));
        Assert.Equal(new[] { "cue1", "cue2" }, CueIds(book));
        book.ReplaceMediaOverlay("overlay", Overlay(Cue("second")));
        Assert.Equal(new[] { "cue1" }, CueIds(book));
        Assert.All(Durations(book), e => Assert.Equal("1s", e.Value));
        book.Write();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IncomingRemovedCueReferencesRejectTheWholeEdit(bool css) {
        var book = Book();
        if (css) book.AddResource("style", "EPUB/style.css", "text/css", System.Text.Encoding.UTF8.GetBytes("p { background-image: url('read.smil#cue0'); }"));
        else {
            var content = book.GetContentXml("chapter");
            content.Root!.Element(Html + "body")!.Add(new XElement(Html + "a", new XAttribute("href", "read.smil#cue0"), "Cue"));
            book.SetContentXml("chapter", content);
        }
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidOperationException>(() => book.ReplaceMediaOverlay("overlay", Overlay(Cue("second"), Cue("third"))));
        Assert.Equal(before, book.Write().Bytes);
        book.ReplaceMediaOverlay("overlay", Overlay(Cue("first", 2), Cue("third", 4)));
        Assert.Equal(new[] { "cue0", "cue2" }, CueIds(book));
    }

    [Fact]
    public void OtherOverlayDurationAndMetadataAttributesSurviveReplacement() {
        var book = Book();
        book.AddChapter("other", "EPUB/other.xhtml", "Other", "<p id='other'>Other.</p>");
        book.AddMediaOverlay("other", "other-overlay", "EPUB/other.smil", Overlay(Cue("other")));
        book.AddMetadataProperty("schema:accessibilitySummary", "Retained summary");
        // Import a publisher's existing metadata identity through the public archive boundary.
        byte[] input = book.Write().Bytes;
        using var buffer = new MemoryStream(); buffer.Write(input, 0, input.Length); buffer.Position = 0;
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, true)) {
            var entry = archive.GetEntry("EPUB/package.opf")!;
            XDocument package;
            using (var stream = entry.Open()) package = XDocument.Load(stream);
            package.Descendants(Opf + "meta").Single(e => (string?)e.Attribute("refines") == "#overlay").SetAttributeValue("id", "narration-duration");
            entry.Delete();
            using var output = archive.CreateEntry("EPUB/package.opf").Open(); package.Save(output);
        }
        book = EpubPublication.Load(new MemoryStream(buffer.ToArray()));
        book.ReplaceMediaOverlay("overlay", Overlay(Cue("first")));
        var metadata = book.GetPackageXml().Descendants(Opf + "meta").ToArray();
        Assert.Equal("1s", metadata.Single(e => (string?)e.Attribute("id") == "narration-duration").Value);
        Assert.Equal("2s", Durations(book).Single(e => e.Attribute("refines") == null).Value);
        Assert.Contains(metadata, e => e.Value == "Retained summary");
    }

    [Theory]
    [InlineData("structure")]
    [InlineData("instruction")]
    [InlineData("association")]
    public void UnsupportedPreservationBoundariesFailWithoutChangingBytes(string boundary) {
        var book = Book();
        if (boundary == "association") {
            book.AddChapter("other", "EPUB/other.xhtml", "Other", "<p>Other.</p>").MediaOverlayId = "overlay";
        } else {
            var smil = book.GetContentXml("overlay");
            if (boundary == "structure") smil.Descendants(Smil + "seq").Single().SetAttributeValue("id", "sequence");
            else smil.AddFirst(new XProcessingInstruction("xml-stylesheet", "href='style.css'"));
            byte[] input = book.Write().Bytes;
            using var buffer = new MemoryStream(); buffer.Write(input, 0, input.Length); buffer.Position = 0;
            using (var zip = new ZipArchive(buffer, ZipArchiveMode.Update, true)) {
                zip.GetEntry("EPUB/read.smil")!.Delete();
                using var output = zip.CreateEntry("EPUB/read.smil").Open(); smil.Save(output);
            }
            book = EpubPublication.Load(new MemoryStream(buffer.ToArray()));
        }
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.ReplaceMediaOverlay("overlay", Overlay(Cue("first"))));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void BudgetCancellationAndInvalidTargetsAreAtomic() {
        var book = Book(); byte[] input = book.Write().Bytes;
        using var archive = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxExpandedBytes = archive.Entries.Sum(e => e.Length) });
        Assert.Throws<InvalidDataException>(() => limited.ReplaceMediaOverlay("overlay", Overlay(Cue("first"), Cue("second"), Cue("third"))));
        Assert.Equal(input, limited.Write().Bytes);
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => book.ReplaceMediaOverlay("overlay", Overlay(Cue("first")), cancel.Token));
        Assert.Throws<InvalidDataException>(() => book.ReplaceMediaOverlay("overlay", Overlay(Cue("missing"))));
        Assert.Equal(input, book.Write().Bytes);
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Narration", "en", "urn:example:replacement");
        book.AddChapter("chapter", "EPUB/chapter.xhtml", "Chapter", "<p id='first'>First.</p><p id='second'>Second.</p><p id='third'>Third.</p>");
        book.AddResource("audio", "EPUB/audio.mp3", "audio/mpeg", new byte[] { 1, 2, 3 });
        book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first"), Cue("second")));
        return book;
    }
    private static EpubMediaOverlayCue Cue(string id, int begin = 0) => new(id, "audio", TimeSpan.FromSeconds(begin), TimeSpan.FromSeconds(begin + 1));
    private static EpubMediaOverlay Overlay(params EpubMediaOverlayCue[] cues) => new() { Cues = cues, AudioDurations = new Dictionary<string, TimeSpan> { ["audio"] = TimeSpan.FromSeconds(10) } };
    private static string?[] CueIds(EpubPublication book) => book.GetContentXml("overlay").Descendants(Smil + "par").Select(e => (string?)e.Attribute("id")).ToArray();
    private static XElement[] Durations(EpubPublication book) => book.GetPackageXml().Descendants(Opf + "meta").Where(e => (string?)e.Attribute("property") == "media:duration").ToArray();
    private static byte[] ReadEntry(byte[] bytes, string path) {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var stream = zip.GetEntry(path)!.Open(); using var output = new MemoryStream(); stream.CopyTo(output); return output.ToArray();
    }
}
