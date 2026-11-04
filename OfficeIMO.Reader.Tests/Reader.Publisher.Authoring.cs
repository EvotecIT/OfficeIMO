using OfficeIMO.Drawing;
using OfficeIMO.Publisher;
using OfficeIMO.Reader.Publisher;
using OfficeIMO.Tests;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderPublisherAuthoringTests {
    [Fact]
    public void SourceFreePublicationProjectsCompleteStoriesAndCurrentTablesWithoutNativeRecoveryClaims() {
        PublisherDocument document = PublisherDocument.Create(792, 612);
        PublisherPage master = document.AddMasterPage("Letterhead");
        document.AddTextFrame(master.Id, 72, 24, 648, 24, Paragraphs("Publication header"));
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        string body = string.Concat(Enumerable.Repeat("An authored story stays complete across fixed pages. ", 25));
        PublisherTextFrame first = document.AddTextFrame(document.Pages[0].Id, 72, 72, 180, 36, Paragraphs(body));
        PublisherPage second = document.AddPage(master.Id);
        document.AddLinkedTextFrame(first.Id, second.Id, 72, 72, 648, 420);
        PublisherTable table = document.AddTable(second.Id, 72, 510, new[] { 120D, 90D }, new[] { 24D });
        document.ReplaceTableCellParagraphs(table.Id, 0, 0, Paragraphs("Created value"));
        OfficeDocumentReadResult result = document.ToOfficeDocumentReadResult("authored.pub", new ReaderOptions { ComputeHashes = true, MaxChars = 256 });
        Assert.Contains("officeimo.reader.publisher", result.CapabilitiesUsed);
        Assert.DoesNotContain("officeimo.publisher.native-recovery", result.CapabilitiesUsed);
        Assert.All(result.Pages, page => Assert.Equal("publisher-authored-page", page.Location!.SourceBlockKind));
        string text = string.Concat(result.Blocks.Select(block => block.Text));
        Assert.Equal(1, text.Split(new[] { body }, StringSplitOptions.None).Length - 1);
        Assert.Equal(1, text.Split(new[] { "Publication header" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(1, text.Split(new[] { "Created value" }, StringSplitOptions.None).Length - 1);
        ReaderTable current = Assert.Single(result.EnumerateTables());
        Assert.Equal("Created value", current.Rows[0][0]);
        Assert.Equal(2, current.Location!.Page);
        Assert.Null(result.Source.SourceHash); Assert.Null(result.Source.LengthBytes);
    }

    [Fact]
    public void AppendedPagesKeepAuthoredLocationsInsideANativeBackedPublication() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Publisher.Tests", "Fixtures",
            "NativeComparison", "TextBlank", "native-textblank.pub");
        PublisherDocument document = PublisherDocument.Load(path);
        document.AddPage();
        OfficeDocumentReadResult result = document.ToOfficeDocumentReadResult("native-with-added-page.pub");
        Assert.Contains("officeimo.publisher.native-recovery", result.CapabilitiesUsed);
        Assert.Equal(new[] { "publisher-page", "publisher-page", "publisher-authored-page" },
            result.Pages.Select(page => page.Location!.SourceBlockKind));
        Assert.Equal(new int?[] { 1, 2, 3 }, result.Pages.Select(page => page.Number));
    }

    private static OfficeRichTextParagraph[] Paragraphs(string text) => new[] {
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, 10, OfficeColor.Black, fontFamily: "Arial") })
    };
}
