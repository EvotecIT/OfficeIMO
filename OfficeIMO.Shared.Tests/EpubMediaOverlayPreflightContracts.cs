using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMediaOverlayPreflightContracts {
    private static readonly XNamespace Smil = "http://www.w3.org/ns/SMIL";

    [Theory]
    [InlineData("text")]
    [InlineData("audio")]
    [InlineData("timing")]
    [InlineData("duration")]
    [InlineData("association")]
    public void LowLevelEditsCannotHideBrokenNarrationFromPreflight(string change) {
        var book = Book();
        if (change == "text") {
            var content = book.GetContentXml("chapter");
            content.Descendants().Single(e => (string?)e.Attribute("id") == "sentence").SetAttributeValue("id", "changed");
            book.SetContentXml("chapter", content);
        } else if (change == "audio") book.RemoveResource("audio");
        else if (change == "duration") book.SetMetadataProperty("media:duration", "9s", "#overlay");
        else if (change == "association") book.Manifest.Single(item => item.Id == "chapter").MediaOverlayId = null;
        else {
            var smil = book.GetContentXml("overlay");
            smil.Descendants(Smil + "audio").Single().SetAttributeValue("clipEnd", "0s");
            book.UpdateResource("overlay", System.Text.Encoding.UTF8.GetBytes(smil.ToString()));
        }
        byte[] before = book.Write().Bytes;
        var preflight = book.Preflight();
        Assert.Contains(preflight.Checks, check => check.Code == "media-overlays" && check.Status == EpubPreflightStatus.Failed);
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Narrated", "en", "urn:example:overlay-preflight");
        book.AddChapter("chapter", "EPUB/chapter.xhtml", "Chapter", "<p id='sentence'>A sentence.</p>");
        book.AddResource("audio", "EPUB/audio.mp3", "audio/mpeg", new byte[] { 1, 2, 3 });
        book.AddMediaOverlay("chapter", "overlay", "EPUB/narration.smil", new EpubMediaOverlay {
            AudioDurations = new Dictionary<string, TimeSpan> { ["audio"] = TimeSpan.FromSeconds(2) },
            Cues = new[] { new EpubMediaOverlayCue("sentence", "audio", TimeSpan.Zero, TimeSpan.FromSeconds(2)) }
        });
        return book;
    }

    [Fact]
    public void ValidClipsPassButDecodingAndPlaybackRemainUnchecked() {
        var result = Book().Preflight();
        Assert.Equal(EpubPreflightStatus.Passed, result.Checks.Single(c => c.Code == "media-overlays").Status);
        Assert.Equal(EpubPreflightStatus.NotChecked, result.Checks.Single(c => c.Code == "media-overlay-audio-decoding").Status);
        Assert.Equal(EpubPreflightStatus.NotChecked, result.Checks.Single(c => c.Code == "reading-system-presentation").Status);
    }

    [Fact]
    public void RetainedNestedSequencesAndTextOnlyCuesAreCheckedWithoutRewriting() {
        var book = Book(); var xml = book.GetContentXml("overlay");
        var sequence = xml.Descendants(Smil + "seq").Single();
        var par = sequence.Element(Smil + "par")!; par.Remove();
        sequence.Add(new XElement(Smil + "seq", new XAttribute(XName.Get("textref", "http://www.idpf.org/2007/ops"), "chapter.xhtml#sentence"), par));
        sequence.Add(new XElement(Smil + "par", new XElement(Smil + "text", new XAttribute("src", "chapter.xhtml#sentence"))));
        book.UpdateResource("overlay", System.Text.Encoding.UTF8.GetBytes(xml.ToString()));
        var imported = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        byte[] before = imported.Write().Bytes;
        Assert.Equal(EpubPreflightStatus.Passed, imported.Preflight().Checks.Single(c => c.Code == "media-overlays").Status);
        Assert.Equal(before, imported.Write().Bytes);
    }

    [Fact]
    public void OmittedClipEndIsAnExplicitUncheckedDurationRatherThanAFalsePass() {
        var book = Book(); var xml = book.GetContentXml("overlay");
        xml.Descendants(Smil + "audio").Single().SetAttributeValue("clipEnd", null);
        book.UpdateResource("overlay", System.Text.Encoding.UTF8.GetBytes(xml.ToString()));
        var check = book.Preflight().Checks.Single(c => c.Code == "media-overlays");
        Assert.Equal(EpubPreflightStatus.NotChecked, check.Status);
        Assert.Contains(check.Diagnostics, d => d.Code == "EPUB_PREFLIGHT_OVERLAY_DURATION_UNCHECKED");
    }

    [Theory]
    [InlineData("total")]
    [InlineData("duplicate")]
    [InlineData("clock")]
    [InlineData("sequence-target")]
    public void MetadataAndContainerReferenceErrorsAreReported(string change) {
        var book = Book();
        if (change == "total") book.SetMetadataProperty("media:duration", "9s");
        else if (change == "duplicate") book.AddMetadataProperty("media:duration", "2s", "#overlay");
        else if (change == "clock") book.SetMetadataProperty("media:duration", "-2s", "#overlay");
        else {
            var xml = book.GetContentXml("overlay");
            xml.Descendants(Smil + "seq").Single().SetAttributeValue(XName.Get("textref", "http://www.idpf.org/2007/ops"), "chapter.xhtml#missing");
            book.UpdateResource("overlay", System.Text.Encoding.UTF8.GetBytes(xml.ToString()));
        }
        Assert.Equal(EpubPreflightStatus.Failed, book.Preflight().Checks.Single(c => c.Code == "media-overlays").Status);
    }

    [Fact]
    public void CancellationIsNotConvertedIntoAValidationFinding() {
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => Book().Preflight(cancellationToken: cancellation.Token));
    }
}
