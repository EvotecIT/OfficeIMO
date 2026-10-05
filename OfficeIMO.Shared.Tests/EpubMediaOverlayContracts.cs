using OfficeIMO.Epub;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMediaOverlayContracts {
    private static readonly XNamespace Smil = "http://www.w3.org/ns/SMIL";
    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";

    [Fact]
    public void OrderedClipsRoundTripWithExactDurationsAndEscapedReferences() {
        var book = Book(); var chapter = book.Manifest.Single(item => item.Id == "chapter");
        book.AddMediaOverlay("chapter", "overlay", "EPUB/overlays/read aloud.smil", Overlay(
            new EpubMediaOverlayCue("first", "audio", TimeSpan.FromTicks(1), TimeSpan.FromTicks(15000001)),
            new EpubMediaOverlayCue("second:é", "audio", TimeSpan.FromSeconds(3), TimeSpan.FromSeconds(5))));
        Assert.Equal("overlay", chapter.MediaOverlayId);
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var smil = reopened.GetContentXml("overlay");
        Assert.Equal(new[] { "../text/chapter%20one.xhtml#first", "../text/chapter%20one.xhtml#second%3A%C3%A9" },
            smil.Descendants(Smil + "text").Select(e => (string?)e.Attribute("src")));
        Assert.Equal("0.0000001s", (string?)smil.Descendants(Smil + "audio").First().Attribute("clipBegin"));
        Assert.Equal("../audio/read%20aloud.mp3", (string?)smil.Descendants(Smil + "audio").First().Attribute("src"));
        Assert.Equal(new[] { "3.5s", "3.5s" }, Durations(reopened).Select(e => e.Value));
        reopened.RenameResource("chapter", "EPUB/new/chapter.xhtml");
        reopened.RenameResource("audio", "EPUB/new/audio.mp3");
        reopened.RenameResource("overlay", "EPUB/new/read.smil");
        var renamed = reopened.GetContentXml("overlay");
        var textReference = EpubReference.Resolve("EPUB/new/read.smil", (string)renamed.Descendants(Smil + "text").First().Attribute("src")!);
        Assert.Equal("EPUB/new/chapter.xhtml", textReference.ContainerPath);
        Assert.Equal("first", textReference.Fragment);
        Assert.Equal("EPUB/new/audio.mp3", EpubReference.Resolve("EPUB/new/read.smil", (string)renamed.Descendants(Smil + "audio").First().Attribute("src")!).ContainerPath);
        reopened.Write();
    }

    [Fact]
    public void AddingASecondChapterRecalculatesTotalFromRetainedOverlayDurations() {
        var book = Book();
        book.AddMediaOverlay("chapter", "first-overlay", "EPUB/first.smil", Overlay(Cue("first")));
        book.SetMetadataProperty("media:duration", "00:00:02", "#first-overlay");
        book.SetMetadataProperty("media:duration", "999s"); // stale total is repaired
        book.AddChapter("next", "EPUB/next.xhtml", "Next", "<p id='next'>Next.</p>");
        book.AddMediaOverlay("next", "next-overlay", "EPUB/next.smil", Overlay(Cue("next")));
        Assert.Equal("3s", Durations(book).Single(e => e.Attribute("refines") == null).Value);
        Assert.Equal("00:00:02", Durations(book).Single(e => (string?)e.Attribute("refines") == "#first-overlay").Value);
    }

    [Theory]
    [InlineData(-1, 1)]
    [InlineData(1, 1)]
    [InlineData(2, 1)]
    [InlineData(0, 11)]
    public void InvalidTimingDoesNotChangeAnyBytes(int begin, int end) {
        var book = Book(); byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil",
            Overlay(new EpubMediaOverlayCue("first", "audio", TimeSpan.FromSeconds(begin), TimeSpan.FromSeconds(end)))));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData("missing", "second:é")]
    [InlineData("second:é", "first")]
    [InlineData("first", "first")]
    [InlineData("section", "first")]
    public void MissingRepeatedReversedOrNestedTargetsAreRejected(string first, string second) {
        var book = Book(); byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue(first), Cue(second))));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void MissingDurationAndUnsupportedAudioAreRejectedWithoutMutation() {
        var book = Book(); byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil",
            new EpubMediaOverlay { Cues = new[] { Cue("first") } }));
        book.Manifest.Single(item => item.Id == "audio").MediaType = "audio/wav";
        var unsupported = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first"))));
        Assert.Equal(unsupported, book.Write().Bytes);
        Assert.DoesNotContain(book.Manifest, item => item.Id == "overlay");
    }

    [Fact]
    public void ExistingAssociationAndMissingRetainedDurationAreNotOverwritten() {
        var book = Book(); book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first")));
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidOperationException>(() => book.AddMediaOverlay("chapter", "other", "EPUB/other.smil", Overlay(Cue("first"))));
        Assert.Equal(before, book.Write().Bytes);
        var imported = Book();
        imported.AddResource("retained", "EPUB/retained.smil", "application/smil+xml", System.Text.Encoding.UTF8.GetBytes("<smil xmlns='http://www.w3.org/ns/SMIL' version='3.0'><body/></smil>"));
        before = imported.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => imported.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first"))));
        Assert.Equal(before, imported.Write().Bytes);
    }

    [Fact]
    public void BudgetAndCancellationFailuresAreAtomic() {
        var book = Book(); byte[] input = book.Write().Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxExpandedBytes = zip.Entries.Sum(e => e.Length) + 50 });
        Assert.Throws<InvalidDataException>(() => limited.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first"))));
        Assert.Equal(input, limited.Write().Bytes);
        var packageLimited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxMetadataBytes = zip.GetEntry("EPUB/package.opf")!.Length + 10 });
        Assert.Throws<InvalidDataException>(() => packageLimited.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first"))));
        Assert.Equal(input, packageLimited.Write().Bytes);
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first")), cancel.Token));
        Assert.Equal(input, book.Write().Bytes);
    }

    [Theory]
    [InlineData("01:02:03.5", 3723.5)]
    [InlineData("02:03.5", 123.5)]
    [InlineData("250ms", 0.25)]
    [InlineData("1.5min", 90)]
    [InlineData("0.5h", 1800)]
    public void RetainedSmilClockFormsAreParsedExactly(string text, double seconds) {
        var book = WithRetainedDuration(text);
        book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first")));
        Assert.Equal(((decimal)seconds + 1).ToString("0.#######", System.Globalization.CultureInfo.InvariantCulture) + "s",
            Durations(book).Single(e => e.Attribute("refines") == null).Value);
    }

    [Theory]
    [InlineData("1e2s")]
    [InlineData("-1s")]
    [InlineData("01:60:00")]
    [InlineData("00:60")]
    [InlineData("0.00000001s")]
    [InlineData("99999999999999999999999999h")]
    public void InvalidOrUnrepresentableRetainedClocksAreRejected(string text) {
        var book = WithRetainedDuration(text); var before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first"))));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication WithRetainedDuration(string duration) {
        var book = Book();
        book.AddResource("retained", "EPUB/retained.smil", "application/smil+xml", System.Text.Encoding.UTF8.GetBytes("<smil xmlns='http://www.w3.org/ns/SMIL' version='3.0'><body/></smil>"));
        book.AddMetadataProperty("media:duration", duration, "#retained");
        return book;
    }

    [Fact]
    public void VersionPrefixAndCueBoundsAreEnforced() {
        var legacy = EpubPublication.Create("Legacy", version: EpubVersion.Epub2);
        Assert.Throws<NotSupportedException>(() => legacy.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first"))));
        var book = Book(); var before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay()));
        Assert.Throws<ArgumentException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Enumerable.Repeat(Cue("first"), 10001).ToArray())));
        Assert.Equal(before, book.Write().Bytes);
        book.DeclareVocabularyPrefix("media", "urn:example:other:"); before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first"))));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void DurationAliasesAreRecognizedAndAmbiguityIsRejected() {
        var book = Book();
        book.DeclareVocabularyPrefix("mo", "http://www.idpf.org/epub/vocab/overlays/#");
        book.AddMetadataProperty("mo:duration", "0s");
        book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first")));
        using var zip = new ZipArchive(new MemoryStream(book.Write().Bytes), ZipArchiveMode.Read);
        using var stream = zip.GetEntry("EPUB/package.opf")!.Open();
        Assert.Equal("1s", XDocument.Load(stream).Descendants(Opf + "meta").Single(e => (string?)e.Attribute("property") == "mo:duration").Value);
        var ambiguous = Book(); ambiguous.AddMetadataProperty("media:duration", "0s"); ambiguous.AddMetadataProperty("media:duration", "1s");
        var before = ambiguous.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => ambiguous.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first"))));
        Assert.Equal(before, ambiguous.Write().Bytes);
    }

    [Fact]
    public void DuplicateContentIdsAndDurationOverflowCannotCreateAnOverlay() {
        var book = Book(); var content = book.GetContentXml("chapter");
        content.Root!.Descendants().First(e => (string?)e.Attribute("id") == "second:é").SetAttributeValue("id", "first");
        book.SetContentXml("chapter", content);
        Assert.Throws<InvalidDataException>(() => book.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", Overlay(Cue("first"))));
        Assert.DoesNotContain(book.Manifest, item => item.Id == "overlay");
        var overflow = Book(); var before = overflow.Write().Bytes;
        var options = new EpubMediaOverlay {
            AudioDurations = new Dictionary<string, TimeSpan> { ["audio"] = TimeSpan.MaxValue },
            Cues = new[] { new EpubMediaOverlayCue("first", "audio", TimeSpan.Zero, TimeSpan.MaxValue), Cue("second:é") }
        };
        Assert.Throws<ArgumentException>(() => overflow.AddMediaOverlay("chapter", "overlay", "EPUB/read.smil", options));
        Assert.Equal(before, overflow.Write().Bytes);
    }

    private static EpubMediaOverlayCue Cue(string id) => new EpubMediaOverlayCue(id, "audio", TimeSpan.Zero, TimeSpan.FromSeconds(1));
    private static EpubMediaOverlay Overlay(params EpubMediaOverlayCue[] cues) => new EpubMediaOverlay {
        Cues = cues, AudioDurations = new Dictionary<string, TimeSpan> { ["audio"] = TimeSpan.FromSeconds(10) }
    };
    private static EpubPublication Book() {
        var book = EpubPublication.Create("Narration", "en", "urn:example:smil");
        book.AddChapter("chapter", "EPUB/text/chapter one.xhtml", "Chapter", "<section id='section'><p id='first'>First.</p><p id='second:é'>Second.</p></section>");
        book.AddResource("audio", "EPUB/audio/read aloud.mp3", "audio/mpeg", new byte[] { 1, 2, 3 });
        return book;
    }
    private static XElement[] Durations(EpubPublication book) {
        using var zip = new ZipArchive(new MemoryStream(book.Write().Bytes), ZipArchiveMode.Read);
        using var stream = zip.GetEntry("EPUB/package.opf")!.Open();
        return XDocument.Load(stream).Descendants(Opf + "meta").Where(e => (string?)e.Attribute("property") == "media:duration").ToArray();
    }
}
