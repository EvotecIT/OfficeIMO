using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Epub;
using Xunit;
using static OfficeIMO.Shared.Tests.EpubIntegrityFixtures;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubReadingFallbackContracts {
    [Theory]
    [InlineData("application/xhtml+xml", true)]
    [InlineData("application/xhtml+xml", false)]
    [InlineData("image/svg+xml", true)]
    [InlineData("image/svg+xml", false)]
    public async Task Load_FollowsForeignChainsAndPreservesRepeatedPositionPolicy(string mediaType, bool spineOrder) {
        byte[] package = CreatePackage(mediaType);
        var options = new EpubReadOptions { PreferSpineOrder = spineOrder, IncludeNonLinearSpineItems = false, IncludeRawHtml = true };
        EpubDocument document = EpubDocument.Load(new MemoryStream(package), options);
        Assert.Equal(new int?[] { 1, 3 }, document.Chapters.Select(chapter => chapter.SpineIndex));
        Assert.All(document.Chapters, chapter => {
            Assert.Equal("EPUB/fallback.xhtml", chapter.Path);
            Assert.Equal("fallback", chapter.ManifestId);
            Assert.Equal(mediaType, chapter.MediaType);
            Assert.Equal("Fallback text", chapter.Text);
            Assert.NotNull(chapter.Html);
        });
        Assert.Equal(2, document.ReadSummary.RequestedChapterCount);
        Assert.True(document.ReadSummary.IsComplete);
        Assert.False(document.ReadSummary.UsedFallbackScan);
        EpubDocument asynchronous = await EpubDocument.LoadAsync(new MemoryStream(package), options);
        Assert.Equal(document.Chapters.Select(chapter => chapter.Path), asynchronous.Chapters.Select(chapter => chapter.Path));
        Assert.Equal(document.Chapters.Select(chapter => chapter.Text), asynchronous.Chapters.Select(chapter => chapter.Text));
    }

    [Theory]
    [InlineData("missing", "epub.spine.fallback-missing")]
    [InlineData("cycle", "epub.spine.fallback-cycle")]
    [InlineData("unsupported", "epub.spine.unsupported-media-type")]
    [InlineData("foreign-attribute", "epub.spine.unsupported-media-type")]
    public void Load_InvalidFallbackChainsRemainDiagnosedAndDoNotTriggerArchiveRecovery(string failure, string code) {
        byte[] package = CreatePackage("application/xhtml+xml", failure);
        EpubDocument document = EpubDocument.Load(new MemoryStream(package));
        Assert.Empty(document.Chapters);
        Assert.Equal(3, document.ReadSummary.RequestedChapterCount);
        Assert.Equal(3, document.ReadSummary.SkippedChapterCount);
        Assert.False(document.ReadSummary.UsedFallbackScan);
        Assert.False(document.ReadSummary.IsComplete);
        Assert.Single(document.Diagnostics, diagnostic => diagnostic.Code == code);
    }

    [Theory]
    [InlineData("count", "epub.chapter.count-limit", 1)]
    [InlineData("bytes", "epub.chapter.size-limit", 0)]
    [InlineData("text", "epub.chapter.text-total-limit", 0)]
    public void Load_FallbacksUseExistingChapterBudgets(string budget, string code, int chapters) {
        var options = new EpubReadOptions {
            IncludeNonLinearSpineItems = false,
            MaxChapters = budget == "count" ? 1 : 500,
            MaxChapterBytes = budget == "bytes" ? 1 : 4096,
            MaxTotalTextCharacters = budget == "text" ? 1 : 4096
        };
        EpubDocument document = EpubDocument.Load(new MemoryStream(CreatePackage("application/xhtml+xml")), options);
        Assert.Equal(chapters, document.Chapters.Count);
        Assert.Equal(2, document.ReadSummary.RequestedChapterCount);
        Assert.Equal(2 - chapters, document.ReadSummary.SkippedChapterCount);
        Assert.Single(document.Diagnostics, diagnostic => diagnostic.Code == code);
    }

    [Fact]
    public void Load_SupportedPrimaryDoesNotPreferItsFallback() {
        byte[] package = EditPackage(CreatePackage("application/xhtml+xml"), root =>
            root.Descendants(Opf + "item").Single(item => (string?)item.Attribute("id") == "foreign").SetAttributeValue("media-type", "application/xhtml+xml"));
        EpubDocument document = EpubDocument.Load(new MemoryStream(package));
        Assert.All(document.Chapters, chapter => Assert.Equal("Primary text", chapter.Text));
        Assert.True(document.ReadSummary.IsComplete);
    }

    internal static byte[] CreatePackage(string mediaType, string? failure = null) {
        byte[] package = Package(new[] {
            ("foreign", "foreign.bin", "application/vnd.example.foreign", ""),
            ("middle", "middle.bin", "application/vnd.example.middle", ""),
            ("fallback", "fallback.xhtml", mediaType, "")
        }, "<itemref idref='foreign'/><itemref idref='foreign' linear='no'/><itemref idref='foreign'/>", new[] {
            ("foreign.bin", Xhtml("<p>Primary text</p>")), ("middle.bin", "opaque"),
            ("fallback.xhtml", mediaType == "image/svg+xml" ? "<svg xmlns='http://www.w3.org/2000/svg'><text>Fallback text</text></svg>" : Xhtml("<p>Fallback text</p>"))
        });
        return EditPackage(package, root => {
            XElement foreign = root.Descendants(Opf + "item").Single(item => (string?)item.Attribute("id") == "foreign");
            XElement middle = root.Descendants(Opf + "item").Single(item => (string?)item.Attribute("id") == "middle");
            foreign.SetAttributeValue(failure == "foreign-attribute" ? XName.Get("fallback", "urn:foreign") : "fallback", "middle");
            middle.SetAttributeValue("fallback", failure == "missing" ? "absent" : failure == "cycle" ? "foreign" : failure == "unsupported" ? null : "fallback");
        });
    }

    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";
    private static byte[] EditPackage(byte[] package, Action<XElement> edit) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using Stream stream = archive.GetEntry("EPUB/package.opf")!.Open();
        XDocument xml = XDocument.Load(stream);
        edit(xml.Root!);
        return ReplaceEntry(package, "EPUB/package.opf", Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)));
    }
}
