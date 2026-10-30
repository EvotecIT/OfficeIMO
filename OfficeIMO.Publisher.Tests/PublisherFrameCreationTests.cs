using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherFrameCreationTests {
    [Fact]
    public void Continuation_creation_shares_one_styled_story_and_reaches_both_page_exports_without_changing_source_evidence() {
        PublisherDocument document = Load();
        PublisherTextFrame original = Assert.Single(document.Pages[0].TextFrames);
        var sourceReport = document.SourceReadReport;
        document.ReplaceStoryParagraphs(original.StoryId, Paragraphs());
        document.SetTextFrameBounds(original.Id, 54, 108, 228, 120);
        PublisherPage oldPage = document.Pages[0];
        PublisherTextFrame oldFrame = Assert.Single(oldPage.TextFrames);
        string before = document.ToSvg();
        Assert.True(oldFrame.HasOverflow);
        int records = document.ReadReport.InspectedRecordCount;
        int originalObjects = document.ReadReport.SourceObjectCount;
        PublisherTextFrame created = document.AddLinkedTextFrame(original.Id, document.Pages[1].Id, 54, 72, 468, 600);
        PublisherTextFrame first = Assert.Single(document.Pages[0].TextFrames);
        Assert.Same(created, Assert.Single(document.Pages[1].TextFrames));
        Assert.NotEqual(original.Id, created.Id); Assert.Equal(original.StoryId, created.StoryId);
        Assert.Equal(created.Id, first.NextFrameId); Assert.Equal(first.Id, created.PreviousFrameId);
        Assert.Null(created.NextFrameId); Assert.Equal(1U, created.Order);
        Assert.Equal(first.ColumnCount, created.ColumnCount); Assert.Equal(first.ColumnSpacing, created.ColumnSpacing);
        Assert.Equal(new OfficePoint(54, 72), created.PageTransform.TransformPoint(new OfficePoint(0, 0)));
        Assert.True(created.TextLength > 0); Assert.False(created.HasOverflow);
        PublisherTextStory story = Assert.Single(document.TextStories);
        Assert.Equal(story.Text, Assigned(document));
        Assert.Contains(story.Paragraphs[0].Runs, run => run.Bold && run.Text == "Continued publication");
        Assert.Contains(story.Paragraphs[1].Runs, run => run.Italic);
        OfficeDrawingRichText[] painted = Painted(document.Pages[1], created.Id);
        Assert.NotEmpty(painted);
        Assert.Contains(painted.SelectMany(region => region.Paragraphs).SelectMany(paragraph => paragraph.Runs), run => run.Italic);
        Assert.DoesNotContain(PublisherNativeTests.Elements(document.Pages[1].CreateDrawing()).OfType<OfficeDrawingShape>(),
            shape => shape.SourceElementIds?.Contains("publisher-object-" + created.Id) == true);
        Assert.Contains("Continuation text", document.ToSvg(1));
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Equal(2, pdf.Pages.Count); Assert.Contains("Continued publication", pdf.ExtractText());
        Assert.Contains("Continuation text", pdf.Pages[1].ExtractText());
        Assert.Equal(before, document.ToSvg());
        Assert.Null(oldFrame.NextFrameId); Assert.True(oldFrame.HasOverflow);
        Assert.Same(sourceReport, document.SourceReadReport);
        Assert.Equal(originalObjects, document.ReadReport.SourceObjectCount);
        Assert.Equal(originalObjects + 1, document.ReadReport.CurrentObjectCount);
        Assert.Equal(originalObjects + 1, document.ReadReport.ProjectedObjectCount);
        Assert.Equal(records, document.ReadReport.InspectedRecordCount);
        Assert.Equal(originalObjects, sourceReport.CurrentObjectCount);
    }

    [Fact]
    public void Authored_frames_remain_editable_and_keep_chain_order_during_subsequent_long_short_and_empty_stories() {
        PublisherDocument document = Load();
        PublisherTextFrame first = Assert.Single(document.Pages[0].TextFrames);
        document.ReplaceStoryParagraphs(first.StoryId, Paragraphs());
        document.SetTextFrameBounds(first.Id, 54, 108, 228, 120);
        PublisherTextFrame second = document.AddLinkedTextFrame(first.Id, document.Pages[1].Id, 54, 72, 228, 120);
        PublisherTextFrame third = document.AddLinkedTextFrame(second.Id, document.Pages[1].Id, 54, 250, 468, 400);
        Assert.Null(second.NextFrameId);
        PublisherTextFrame[] chain = Frames(document);
        Assert.Equal(new[] { first.Id, second.Id, third.Id }, chain.Select(frame => frame.Id));
        Assert.Equal(new[] { first.Id, second.Id }, chain.Skip(1).Select(frame => frame.PreviousFrameId!.Value));
        Assert.Equal(Assigned(document), Assert.Single(document.TextStories).Text);
        document.SetTextFrameBounds(third.Id, -12, 250, 480, 420);
        Assert.Equal(54, third.X);
        Assert.Equal(-12, Frames(document)[2].X);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_PAGE_BLEED_CLIPPED"
            && diagnostic.Location == "publication/frame/" + third.Id);
        document.ReplaceStoryParagraphs(first.StoryId, new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Short", 11, OfficeColor.Black, fontFamily: "Arial") })
        });
        Assert.Equal("Short\n", Assigned(document));
        Assert.All(Frames(document).Skip(1), frame => { Assert.Equal(0, frame.TextLength); Assert.False(frame.HasOverflow); });
        document.ReplaceStoryParagraphs(first.StoryId, Array.Empty<OfficeRichTextParagraph>());
        Assert.Equal(string.Empty, Assert.Single(document.TextStories).Text);
        Assert.All(Frames(document), frame => Assert.Equal(0, frame.TextLength));
        Assert.Equal(new[] { second.Id, third.Id }, document.Pages[1].TextFrames.Select(frame => frame.Id));
    }

    [Fact]
    public void Invalid_tail_page_geometry_and_cancellation_do_not_commit_or_consume_the_next_identity() {
        PublisherDocument document = Load();
        PublisherTextFrame first = Assert.Single(document.Pages[0].TextFrames);
        uint destination = document.Pages[1].Id;
        var pages = document.Pages; var stories = document.TextStories; var report = document.ReadReport;
        string before = document.ToSvg();
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddLinkedTextFrame(uint.MaxValue, destination, 72, 72, 100, 100));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddLinkedTextFrame(first.Id, uint.MaxValue, 72, 72, 100, 100));
        Assert.Throws<NotSupportedException>(() => document.AddLinkedTextFrame(first.Id, document.MasterPages[0].Id, 72, 72, 100, 100));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddLinkedTextFrame(first.Id, destination, double.NaN, 72, 100, 100));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddLinkedTextFrame(first.Id, destination, 72, 72, 0, 100));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddLinkedTextFrame(first.Id, destination, double.MaxValue, 72, 100, 100));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.AddLinkedTextFrame(first.Id, destination, 72, 72, 100, 100, cancelled.Token));
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories); Assert.Same(report, document.ReadReport);
        Assert.Equal(before, document.ToSvg());
        PublisherDocument control = Load();
        PublisherTextFrame expected = control.AddLinkedTextFrame(first.Id, destination, 72, 72, 100, 100);
        PublisherTextFrame actual = document.AddLinkedTextFrame(first.Id, destination, 72, 72, 100, 100);
        Assert.Equal(expected.Id, actual.Id);
        var committed = document.Pages;
        Assert.Throws<NotSupportedException>(() => document.AddLinkedTextFrame(first.Id, destination, 72, 200, 100, 100));
        Assert.Same(committed, document.Pages);
    }

    [Fact]
    public void Master_grouped_table_and_unresolved_chains_are_not_continuation_sources() {
        PublisherDocument masters = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/MultipleMasters/native-multiplemasters.pub"));
        PublisherTextFrame master = masters.MasterPages.SelectMany(page => page.TextFrames).First();
        var masterPages = masters.Pages;
        Assert.Throws<NotSupportedException>(() => masters.AddLinkedTextFrame(master.Id, masters.Pages[0].Id, 72, 72, 100, 100));
        Assert.Same(masterPages, masters.Pages);
        PublisherDocument grouped = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        Assert.Throws<NotSupportedException>(() => grouped.AddLinkedTextFrame(345, grouped.Pages[0].Id, 72, 72, 100, 100));
        PublisherDocument table = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TableLabel/native-tablelabel.pub"));
        uint tableId = table.Pages.SelectMany(page => page.Tables).First().Id;
        Assert.Throws<ArgumentOutOfRangeException>(() => table.AddLinkedTextFrame(tableId, table.Pages[0].Id, 72, 72, 100, 100));
        byte[] unresolved = PublisherInputContractTests.Mutate("Contents", bytes => {
            var data = new OfficeIMO.Publisher.Internal.PublisherBinaryData(bytes, "Contents");
            var context = new OfficeIMO.Publisher.Internal.PublisherParseContext(new PublisherReadOptions(), default);
            var blocks = new OfficeIMO.Publisher.Internal.PublisherBlockReader(data, context);
            var publication = new OfficeIMO.Publisher.Internal.PublisherContentsReader(data, context).Read();
            var document = publication.Chunks.Values.Single(chunk => chunk.Kind == 0x44);
            var catalog = OfficeIMO.Publisher.Internal.PublisherContentsReader.Field(blocks.Chunk(document.Offset, document.End), 0x20)!.Value;
            bytes[catalog.Offset - 2] = 0xFE;
        }, "NativeComparison/Newsletter/native-newsletter.pub");
        PublisherDocument unlinked = PublisherDocument.Load(unresolved);
        var pages = unlinked.Pages;
        Assert.Throws<NotSupportedException>(() => unlinked.AddLinkedTextFrame(298, unlinked.Pages[1].Id, 72, 72, 100, 100));
        Assert.Same(pages, unlinked.Pages);
    }

    [Fact]
    public void Mixed_image_and_text_views_share_the_current_frame_placement_and_keep_old_snapshots_detached() {
        PublisherDocument document = PublisherFrameCreationFixture.Load(mixedPicture: true).Document;
        PublisherTextFrame beforeFrame = Assert.Single(document.Pages[0].TextFrames);
        PublisherPicture beforePicture = Assert.Single(document.Pages[0].Pictures);
        Assert.Equal(beforeFrame.Id, beforePicture.Id); Assert.False(beforePicture.CanEdit);
        string text = Assert.Single(document.TextStories).Text;
        document.SetTextFrameBounds(beforeFrame.Id, 100, 110, 200, 80);
        PublisherTextFrame frame = Assert.Single(document.Pages[0].TextFrames);
        PublisherPicture picture = Assert.Single(document.Pages[0].Pictures);
        Assert.Equal((frame.X, frame.Y, frame.Width, frame.Height), (picture.X, picture.Y, picture.Width, picture.Height));
        Assert.Equal(frame.PageTransform, picture.PageTransform); Assert.Equal(beforeFrame.Id, picture.Id);
        Assert.Equal((100D, 110D, 200D, 80D), (picture.X, picture.Y, picture.Width, picture.Height));
        Assert.Equal(72, beforeFrame.X, 5); Assert.Equal(72, beforePicture.X, 5);
        Assert.Equal(text, Assert.Single(document.TextStories).Text);
        OfficeDrawingImage image = Assert.Single(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingImage>());
        Assert.Equal(100, image.Projection.Placement.X); Assert.Equal(110, image.Projection.Placement.Y);
        Assert.Equal(200, image.Projection.Placement.Width); Assert.Equal(80, image.Projection.Placement.Height);
        Assert.False(picture.CanEdit);
        Assert.Throws<NotSupportedException>(() => document.ReplacePictureImage(picture.Id, PublisherPictureEditingFixture.Image(OfficeColor.Black)));
        Assert.Throws<NotSupportedException>(() => document.AddLinkedTextFrame(frame.Id, document.Pages[1].Id, 72, 72, 100, 100));
    }

    [Fact]
    public void Authored_identity_does_not_resolve_a_preexisting_dangling_wrap_reference() {
        var fixture = PublisherFrameCreationFixture.Load(danglingReference: true);
        PublisherDocument document = fixture.Document;
        PublisherTextFrame first = Assert.Single(document.Pages[0].TextFrames);
        PublisherTextFrame added = document.AddLinkedTextFrame(first.Id, document.Pages[0].Id, 72, 200, 360, 72);
        Assert.NotEqual(fixture.ReservedId, added.Id);
        Assert.Equal(new[] { fixture.ReservedId }, Assert.Single(document.Pages[0].TextFrames, frame => frame.Id == first.Id).WrapObjectIds);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_TEXT_WRAP_REFERENCE_UNRESOLVED"
            && diagnostic.Location == "Contents/object/" + first.Id);
        Assert.Equal(2, document.Pages[0].TextFrames.Count);
    }

    [Fact]
    public void Model_object_and_projection_limits_leave_the_previous_publication_usable() {
        const int itemLimit = 128;
        // TextBlank contains 36 native default-style pairs. Keep its normal
        // admission valid while filling the current object budget explicitly.
        PublisherDocument full = PublisherFrameCreationFixture.Load(extraHiddenObjects: itemLimit - 1,
            options: new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = itemLimit } }).Document;
        PublisherTextFrame frame = Assert.Single(full.Pages[0].TextFrames);
        Assert.Equal(itemLimit, full.ReadReport.CurrentObjectCount);
        var fullPages = full.Pages;
        Assert.Throws<InvalidDataException>(() => full.AddLinkedTextFrame(frame.Id, full.Pages[1].Id, 72, 72, 100, 100));
        Assert.Same(fullPages, full.Pages);

        var boundedOptions = new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = itemLimit, MaxTextCharacters = 512 } };
        PublisherDocument bounded = Load(boundedOptions);
        PublisherTextFrame first = Assert.Single(bounded.Pages[0].TextFrames);
        bounded.SetTextFrameBounds(first.Id, 72, 72, 40, 32);
        // The complete 432-character paragraph is measured once by the narrow
        // original frame. Measuring its remainder in another frame exceeds the
        // same operation's 512-character work ceiling, without changing limits.
        var runs = Enumerable.Range(0, 16).Select(index => new OfficeRichTextRun("Bounded continuation text. ", 11,
            OfficeColor.Black, bold: index % 2 == 0, fontFamily: "Arial")).ToArray();
        bounded.ReplaceStoryParagraphs(first.StoryId, new[] { new OfficeRichTextParagraph(runs) });
        Assert.True(Assert.Single(bounded.Pages[0].TextFrames).HasOverflow);
        var pages = bounded.Pages; var report = bounded.ReadReport; string before = bounded.ToSvg();
        var failure = Assert.Throws<InvalidDataException>(() => bounded.AddLinkedTextFrame(first.Id, bounded.Pages[1].Id, 72, 72, 100, 100));
        Assert.Contains("text layout character work limit", failure.Message);
        Assert.Same(pages, bounded.Pages); Assert.Same(report, bounded.ReadReport); Assert.Equal(before, bounded.ToSvg());
        bounded.ReplaceStoryParagraphs(first.StoryId, new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Short", 11, OfficeColor.Black, fontFamily: "Arial") })
        });
        PublisherTextFrame emptyContinuation = bounded.AddLinkedTextFrame(first.Id, bounded.Pages[1].Id, 72, 72, 100, 100);
        Assert.Equal(0, emptyContinuation.TextLength);
        Assert.Equal(2, bounded.ReadReport.CurrentObjectCount);
        PublisherDocument control = Load(boundedOptions);
        PublisherTextFrame controlFrame = Assert.Single(control.Pages[0].TextFrames);
        Assert.Equal(control.AddLinkedTextFrame(controlFrame.Id, control.Pages[1].Id, 72, 72, 100, 100).Id, emptyContinuation.Id);
    }

    private static PublisherDocument Load(PublisherReadOptions? options = null) =>
        PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TextBlank/native-textblank.pub"), options);
    private static OfficeRichTextParagraph[] Paragraphs() => new[] {
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Continued publication", 16, OfficeColor.Black, bold: true, fontFamily: "Arial") }),
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(string.Concat(Enumerable.Repeat("Continuation text stays complete. ", 60)), 11,
            OfficeColor.Black, italic: true, fontFamily: "Arial") })
    };
    private static PublisherTextFrame[] Frames(PublisherDocument document) => document.Pages.SelectMany(page => page.TextFrames).OrderBy(frame => frame.Order).ToArray();
    private static string Assigned(PublisherDocument document) {
        PublisherTextStory story = Assert.Single(document.TextStories);
        return string.Concat(Frames(document).Select(frame => story.Text.Substring(frame.TextStart!.Value, frame.TextLength!.Value)));
    }
    private static OfficeDrawingRichText[] Painted(PublisherPage page, uint id) => PublisherNativeTests.Elements(page.CreateDrawing())
        .OfType<OfficeDrawingRichText>().Where(text => text.SourceElementIds?.Contains("publisher-object-" + id) == true).ToArray();
}
