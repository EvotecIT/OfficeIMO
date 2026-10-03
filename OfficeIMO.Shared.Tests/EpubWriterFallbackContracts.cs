using System.Text;
using OfficeIMO.Epub;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubWriterFallbackContracts {
    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void ForeignSpine_ProjectsTheSupportedFallbackThroughReadAndReload(EpubVersion version) {
        EpubPublication book = CreateForeignBook(version);
        EpubDocument read = book.Read();
        Assert.Equal("Readable fallback.", Assert.Single(read.Chapters).Text);
        Assert.Equal("EPUB/fallback.xhtml", read.Chapters[0].Path);
        Assert.Equal("fallback", read.Chapters[0].ManifestId);
        Assert.True(read.ReadSummary.IsComplete);
        book = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        book.Title = "Edited fallback book";
        Assert.Equal("Readable fallback.", Assert.Single(book.Read().Chapters).Text);
        Assert.Equal("foreign", Assert.Single(book.Spine).ManifestId);
    }

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void Navigation_AcceptsFallbackSpineTargetsAndRejectsUnrelatedResources(EpubVersion version) {
        EpubPublication book = CreateForeignBook(version);
        book.SetNavigation(new[] { new EpubNavigationEntry("Readable fallback", "EPUB/fallback.xhtml#start") },
            new[] { new EpubNavigationEntry("1", "EPUB/fallback.xhtml#start") },
            new[] { new EpubNavigationEntry("Start", "EPUB/fallback.xhtml#start", semanticType: version == EpubVersion.Epub2 ? "text" : "bodymatter") });
        EpubDocument read = book.Read();
        Assert.Equal("EPUB/fallback.xhtml", Assert.Single(read.TableOfContents).Target);
        Assert.Single(read.PageList);
        Assert.Single(read.Landmarks);
        book.AddResource("other", "EPUB/other.xhtml", "application/xhtml+xml", Encoding.UTF8.GetBytes(EpubIntegrityFixtures.Xhtml("<p>Other</p>")));
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.SetNavigation(new[] { new EpubNavigationEntry("Other", "EPUB/other.xhtml") }));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication CreateForeignBook(EpubVersion version) {
        EpubPublication book = EpubPublication.Create("Fallback book", identifier: "urn:book:fallback", version: version);
        book.AddChapter("fallback", "EPUB/fallback.xhtml", "Fallback chapter", "<p id='start'>Readable fallback.</p>");
        EpubManifestItem foreign = book.AddResource("foreign", "EPUB/foreign.bin", "application/vnd.example.foreign", new byte[] { 1 });
        foreign.FallbackId = "fallback";
        book.RemoveSpineItem(0);
        book.AddSpineItem("foreign");
        book.SetNavigation(new[] { new EpubNavigationEntry("Foreign document", "EPUB/foreign.bin") });
        return book;
    }
}
