using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherMasterPairAuthoringTests {
    [Fact]
    public void Authored_pair_keeps_distinct_side_registered_definitions_and_explicit_assignments() {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        var originalPages = document.Pages; var originalMasters = document.MasterPages;
        PublisherPage single = document.AddMasterPage("A");
        var pair = document.AddMasterPagePair("P");
        Assert.Equal(PublisherMasterPageSide.Single, single.MasterSide);
        Assert.Equal(PublisherMasterPageSide.Left, pair.Left.MasterSide);
        Assert.Equal(PublisherMasterPageSide.Right, pair.Right.MasterSide);
        Assert.Equal(pair.Left.Id + 1, pair.Right.Id);
        Assert.Equal(new[] { single.Id, pair.Left.Id, pair.Right.Id }, document.MasterPages.Select(page => page.Id));
        Assert.All(document.MasterPages, page => { Assert.False(page.HasNativeOrigin); Assert.Null(page.MasterPageId); });
        Assert.All(document.Pages, page => { Assert.Null(page.MasterSide); Assert.Null(page.MasterPageId); });
        PublisherTextFrame left = document.AddTextFrame(pair.Left.Id, 18, 24, 252, 24, PublisherAuthoringTests.Paragraphs("Left definition"));
        PublisherTextFrame right = document.AddTextFrame(pair.Right.Id, 342, 24, 252, 24, PublisherAuthoringTests.Paragraphs("Right definition"));
        document.SetPageMaster(document.Pages[0].Id, pair.Right.Id);
        PublisherPage added = document.AddPage(pair.Left.Id);
        Assert.Empty(added.TextFrames);
        Assert.Equal(324, Marker(added, left.Id).X);
        Assert.Equal(36, Marker(document.Pages[0], right.Id).X);
        Assert.Equal(18, Marker(document.MasterPages.Single(page => page.Id == pair.Left.Id), left.Id).X);
        Assert.Equal(342, Marker(document.MasterPages.Single(page => page.Id == pair.Right.Id), right.Id).X);
        Assert.Empty(pair.Left.TextFrames); Assert.Empty(pair.Right.TextFrames);
        Assert.Empty(originalMasters); Assert.Null(originalPages[0].MasterPageId);
        document.SetPageMaster(added.Id, null);
        Assert.DoesNotContain("Left definition", document.ToSvg(1));
        Assert.Contains("Left definition", PublisherAuthoringTests.PaintedText(added));
        Assert.Equal(2, document.TextStories.Count);
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Contains("Right definition", pdf.Pages[0].ExtractText());
        Assert.DoesNotContain("Left definition", pdf.Pages[1].ExtractText());
    }

    [Fact]
    public void Pair_limits_invalid_names_and_cancellation_leave_both_definitions_and_allocator_unchanged() {
        PublisherDocument limited = PublisherDocument.Create(612, 792, new PublisherReadOptions { MaximumPages = 2 });
        var pages = limited.Pages; var masters = limited.MasterPages; var report = limited.ReadReport;
        Assert.Throws<InvalidDataException>(() => limited.AddMasterPagePair("P"));
        Assert.Same(pages, limited.Pages); Assert.Same(masters, limited.MasterPages); Assert.Same(report, limited.ReadReport);
        Assert.Equal(2U, limited.AddMasterPage("Only free definition").Id);

        PublisherDocument nameLimited = PublisherDocument.Create(612, 792,
            new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 3 } });
        var beforeNames = nameLimited.MasterPages;
        Assert.Throws<ArgumentException>(() => nameLimited.AddMasterPagePair("PP"));
        Assert.Same(beforeNames, nameLimited.MasterPages);
        Assert.Equal(2U, nameLimited.AddMasterPage("One").Id);

        PublisherDocument document = PublisherDocument.Create(612, 792);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.AddMasterPagePair("P", cancellation.Token));
        Assert.Throws<ArgumentNullException>(() => document.AddMasterPagePair(null!));
        Assert.Throws<ArgumentException>(() => document.AddMasterPagePair(" "));
        Assert.Throws<ArgumentException>(() => document.AddMasterPagePair("bad\nname"));
        Assert.Throws<ArgumentException>(() => document.AddMasterPagePair("\uD800"));
        Assert.Empty(document.MasterPages);
        var pair = document.AddMasterPagePair("P");
        Assert.Equal(2U, pair.Left.Id); Assert.Equal(3U, pair.Right.Id);

        PublisherDocument exhausted = PublisherTableEditingFixture.Load("one-id");
        var before = exhausted.MasterPages;
        Assert.Throws<InvalidOperationException>(() => exhausted.AddMasterPagePair("P"));
        Assert.Same(before, exhausted.MasterPages);
        Assert.Equal(uint.MaxValue, exhausted.AddMasterPage("Last identifier").Id);
    }

    [Fact]
    public void Independent_native_master_sides_survive_content_edits_and_native_source_saving() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TwoPageMaster/native-twopagemaster.pub"));
        PublisherMasterPageSide?[] sides = { PublisherMasterPageSide.Single, PublisherMasterPageSide.Left, PublisherMasterPageSide.Right };
        Assert.Equal(sides, document.MasterPages.Select(page => page.MasterSide));
        PublisherTextFrame frame = document.MasterPages[1].TextFrames[0];
        document.SetTextFrameBounds(frame.Id, frame.X + 6, frame.Y + 12, frame.Width, frame.Height);
        PublisherDocument reopened = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Equal(sides, reopened.MasterPages.Select(page => page.MasterSide));
        Assert.Equal(document.Pages.Select(page => page.MasterPageId), reopened.Pages.Select(page => page.MasterPageId));
        Assert.Equal(frame.X + 6, reopened.MasterPages[1].TextFrames[0].X);
        Assert.All(reopened.Pages, page => Assert.Null(page.MasterSide));
    }

    private static OfficeDrawingRichText Marker(PublisherPage page, uint id) =>
        Assert.Single(PublisherNativeTests.Elements(page.CreateDrawing()).OfType<OfficeDrawingRichText>(), item => item.SourceElementIds?.Contains("publisher-object-" + id) == true);
}
