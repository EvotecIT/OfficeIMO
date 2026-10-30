using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherAuthoringTests {
    [Fact]
    public void Source_free_newsletter_keeps_complete_styled_stories_master_ownership_tables_and_exported_pages() {
        PublisherDocument document = PublisherDocument.Create(792, 612);
        PublisherPage originalPage = Assert.Single(document.Pages);
        PublisherReadReport source = document.SourceReadReport;
        Assert.False(document.HasNativeSource); Assert.False(originalPage.HasNativeOrigin);
        Assert.Empty(document.MasterPages); Assert.Empty(document.TextStories); Assert.Empty(document.Images);
        Assert.Equal(0, document.ReadReport.CurrentObjectCount); Assert.Empty(source.FidelityDiagnostics);
        PublisherPage master = document.AddMasterPage("Newsletter header");
        PublisherTextFrame heading = document.AddTextFrame(master.Id, 36, 18, 600, 28, Paragraphs("Weekly publication", bold: true));
        PublisherShape rule = document.AddShape(master.Id, OfficeShape.Rectangle(720, 2), 36, 54);
        document.SetPageMaster(originalPage.Id, master.Id);
        var supplied = new List<OfficeRichTextParagraph> {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Independent styled story", 15, OfficeColor.Black, bold: true, fontFamily: "Arial") }),
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(string.Concat(Enumerable.Repeat("Continued body content. ", 45)), 11,
                OfficeColor.Black, italic: true, fontFamily: "Arial") })
        };
        PublisherTextFrame first = document.AddTextFrame(originalPage.Id, 36, 84, 240, 100, supplied);
        supplied.Clear();
        Assert.True(first.HasOverflow);
        PublisherPage firstSnapshot = document.Pages[0];
        PublisherPage second = document.AddPage(master.Id);
        PublisherTextFrame continuation = document.AddLinkedTextFrame(first.Id, second.Id, 36, 84, 720, 350);
        PublisherTable table = document.AddTable(second.Id, 36, 480, new[] { 120D, 160D }, new[] { 28D });
        byte[] suppliedPicture = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/Image/native-image.png"));
        byte[] expectedPicture = (byte[])suppliedPicture.Clone();
        PublisherPicture illustration = document.AddPicture(originalPage.Id, suppliedPicture, 360, 84, 96, 96);
        Array.Clear(suppliedPicture, 0, suppliedPicture.Length);
        byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(4, 3, OfficeColor.FromRgb(20, 130, 40)));
        PublisherPicture logo = document.AddPicture(master.Id, jpeg, 720, 12, 32, 32);
        document.ReplaceTableCellParagraphs(table.Id, 0, 0, Paragraphs("Table left"));
        document.ReplaceTableCellParagraphs(table.Id, 0, 1, Paragraphs("Table right"));

        Assert.False(continuation.HasOverflow); Assert.True(continuation.TextLength > 0);
        PublisherTextStory body = document.TextStories.Single(story => story.Id == first.StoryId);
        PublisherTextFrame[] frames = document.Pages.SelectMany(page => page.TextFrames).Where(frame => frame.StoryId == body.Id).OrderBy(frame => frame.Order).ToArray();
        Assert.Equal(body.Text, string.Concat(frames.Select(frame => body.Text.Substring(frame.TextStart!.Value, frame.TextLength!.Value))));
        Assert.False(body.Text.EndsWith("\n", StringComparison.Ordinal));
        Assert.True(body.Paragraphs[0].Runs[0].Bold); Assert.True(body.Paragraphs[1].Runs[0].Italic);
        Assert.Single(document.MasterPages[0].TextFrames); Assert.Single(document.MasterPages[0].Shapes);
        Assert.Equal(logo.Id, Assert.Single(document.MasterPages[0].Pictures).Id);
        Assert.Equal(expectedPicture, document.Images.Single(image => image.Id == illustration.ImageId).GetBytes());
        Assert.Equal("image/jpeg", document.Images.Single(image => image.Id == logo.ImageId).ContentType);
        Assert.All(document.Pages, page => {
            Assert.Equal(master.Id, page.MasterPageId); Assert.False(page.HasNativeOrigin);
            Assert.Equal(792, page.Width); Assert.Equal(612, page.Height);
            Assert.DoesNotContain(page.TextFrames, frame => frame.Id == heading.Id);
            Assert.DoesNotContain(page.Shapes, shape => shape.Id == rule.Id);
            Assert.DoesNotContain(page.Pictures, picture => picture.Id == logo.Id);
            Assert.Contains("Weekly publication", PaintedText(page));
        });
        uint[] identities = document.Pages.Select(page => page.Id).Concat(document.MasterPages.Select(page => page.Id))
            .Concat(document.TextStories.Select(story => story.Id)).Concat(document.Pages.Concat(document.MasterPages).SelectMany(page => page.TextFrames).Select(frame => frame.Id))
            .Concat(document.Pages.Concat(document.MasterPages).SelectMany(page => page.Shapes).Select(shape => shape.Id))
            .Concat(document.Pages.Concat(document.MasterPages).SelectMany(page => page.Pictures).Select(picture => picture.Id))
            .Concat(document.Pages.SelectMany(page => page.Tables).Select(item => item.Id)).ToArray();
        Assert.Equal(identities.Length, identities.Distinct().Count());
        Assert.Equal(first.Id + 1, first.StoryId); Assert.Equal(table.Id + 1, table.StoryId);
        Assert.Same(source, document.SourceReadReport); Assert.Equal(0, source.SourceObjectCount);
        Assert.Equal(0, source.TextStoryCount); Assert.Equal(0, source.InspectedRecordCount);
        Assert.Equal(0, document.ReadReport.SourceObjectCount); Assert.Equal(0, document.ReadReport.InspectedRecordCount);
        Assert.Equal(7, document.ReadReport.CurrentObjectCount); Assert.Equal(7, document.ReadReport.ProjectedObjectCount); Assert.Equal(2, document.ReadReport.ImageCount);
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_FIELDS_UNASSESSED");
        Assert.Null(originalPage.MasterPageId); Assert.Empty(originalPage.TextFrames);
        Assert.True(Assert.Single(firstSnapshot.TextFrames).HasOverflow); Assert.Null(Assert.Single(firstSnapshot.TextFrames).NextFrameId);
        Assert.Empty(second.TextFrames); Assert.Empty(second.Tables);
        Assert.Contains("Independent styled story", document.ToSvg()); Assert.Contains("Table right", document.ToSvg(1));
        Assert.Contains(Convert.ToBase64String(expectedPicture), document.ToSvg()); Assert.Contains(Convert.ToBase64String(jpeg), document.ToSvg(1));
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Equal(2, pdf.Pages.Count);
        Assert.All(pdf.Pages, page => { Assert.Equal(792, page.GetPageSize().Width); Assert.Equal(612, page.GetPageSize().Height); Assert.Contains("Weekly publication", page.ExtractText()); });
        Assert.Contains("Continued body content", pdf.Pages[1].ExtractText()); Assert.Contains("Table left", pdf.Pages[1].ExtractText());
        Assert.Equal(new[] { 2, 1 }, pdf.Pages.Select(page => page.GetImagePlacements().Count));
        PdfImagePlacement physicalPicture = Assert.Single(pdf.Pages[0].GetImagePlacements(), image => Math.Abs(image.X - 360) < .01);
        Assert.Equal(96, physicalPicture.Width, 4); Assert.Equal(96, physicalPicture.Height, 4);
    }

    [Fact]
    public void Independent_frames_use_explicit_bounds_zero_insets_and_keep_authored_text_convention_after_edits() {
        PublisherDocument document = PublisherDocument.Create(300, 400);
        uint pageId = document.Pages[0].Id;
        var tabs = new OfficeTextTabStops(new[] { new OfficeTextTabStop(40) });
        var label = OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun("1.", 11, OfficeColor.Black, fontFamily: "Arial"), 0);
        PublisherTextFrame frame = document.AddTextFrame(pageId, 20, 30, 240, 100, new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("First\r\nSecond\rThird\tvalue", 11, OfficeColor.Black, italic: true, fontFamily: "Arial") }, label).WithTabStops(tabs)
        });
        Assert.Equal(1, frame.ColumnCount); Assert.Equal(0, frame.ColumnSpacing); Assert.Empty(frame.WrapObjectIds);
        Assert.Null(frame.PreviousFrameId); Assert.Null(frame.NextFrameId); Assert.Equal(0U, frame.Order);
        OfficeDrawingRichText painted = Assert.Single(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingRichText>());
        Assert.Equal(20, painted.X); Assert.Equal(30, painted.Y); Assert.Equal(240, painted.Width); Assert.Equal(100, painted.Height);
        Assert.DoesNotContain(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingShape>(),
            shape => shape.SourceElementIds?.Contains("publisher-object-" + frame.Id) == true);
        PublisherTextStory story = Assert.Single(document.TextStories);
        Assert.Equal("First\nSecond\nThird\tvalue", story.Text); Assert.Equal("1.", story.Paragraphs[0].Label!.Run.Text);
        Assert.Equal(40, story.Paragraphs[0].TabStops!.Stops[0].Position);
        document.SetTextFrameBounds(frame.Id, -4, 40, 250, 110);
        document.ReplaceStoryParagraphs(frame.StoryId, Array.Empty<OfficeRichTextParagraph>());
        Assert.Empty(Assert.Single(document.TextStories).Text);
        document.ReplaceStoryParagraphs(frame.StoryId, Paragraphs("Repopulated"));
        Assert.Equal("Repopulated", Assert.Single(document.TextStories).Text);
        Assert.Equal(-4, Assert.Single(document.Pages[0].TextFrames).X); Assert.Equal(20, frame.X);
        Assert.Equal("First\nSecond\nThird\tvalue", story.Text);
        PublisherTextFrame empty = document.AddTextFrame(pageId, 100, 200, 40, 20, Array.Empty<OfficeRichTextParagraph>());
        Assert.Equal(0, empty.TextLength); Assert.False(empty.HasOverflow); Assert.NotEqual(frame.StoryId, empty.StoryId);
        document.ReplaceStoryParagraphs(empty.StoryId, Paragraphs("New content"));
        Assert.Equal("New content", document.TextStories.Single(item => item.Id == empty.StoryId).Text);
    }

    [Fact]
    public void Invalid_authoring_and_cancellation_preserve_snapshots_and_both_frame_story_identifiers() {
        PublisherDocument document = PublisherDocument.Create(300, 400);
        uint id = document.Pages[0].Id;
        var pages = document.Pages; var stories = document.TextStories; var report = document.ReadReport;
        string before = document.ToSvg();
        Assert.Throws<ArgumentOutOfRangeException>(() => PublisherDocument.Create(0, 400));
        Assert.Throws<ArgumentOutOfRangeException>(() => PublisherDocument.Create(300, double.PositiveInfinity));
        Assert.Throws<ArgumentOutOfRangeException>(() => PublisherDocument.Create(double.NaN, 400));
        Assert.Throws<ArgumentOutOfRangeException>(() => PublisherDocument.Create(300, 400, new PublisherReadOptions { MaximumPages = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTextFrame(900, 0, 0, 100, 20, Paragraphs("Invalid owner")));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTextFrame(id, double.NaN, 0, 100, 20, Paragraphs("Invalid point")));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTextFrame(id, 0, 0, 0, 20, Paragraphs("Invalid size")));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTextFrame(id, double.MaxValue, 0, double.MaxValue, 20, Paragraphs("Overflowing extent")));
        Assert.Throws<ArgumentException>(() => document.AddTextFrame(id, 0, 0, 100, 20, Paragraphs("\uD800")));
        Assert.Throws<ArgumentException>(() => document.AddTextFrame(id, 0, 0, 100, 20, Paragraphs("bad\0text")));
        Assert.Throws<ArgumentException>(() => document.AddMasterPage("\uD800"));
        Assert.Throws<ArgumentException>(() => document.AddMasterPage("\n"));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddPage(900));
        Assert.Throws<NotSupportedException>(() => document.SetPageMaster(id, id));
        var cancelled = new CancellationToken(true);
        Assert.Throws<OperationCanceledException>(() => document.AddTextFrame(id, 0, 0, 100, 20, Paragraphs("Cancelled"), cancelled));
        Assert.Throws<OperationCanceledException>(() => document.AddPage(null, cancelled));
        Assert.Throws<OperationCanceledException>(() => document.AddMasterPage("Cancelled", cancelled));
        Assert.Throws<OperationCanceledException>(() => document.SetPageMaster(id, null, cancelled));
        Assert.Throws<OperationCanceledException>(() => PublisherDocument.Create(300, 400, cancellationToken: cancelled));
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories); Assert.Same(report, document.ReadReport); Assert.Equal(before, document.ToSvg());
        PublisherTextFrame created = document.AddTextFrame(id, 20, 20, 100, 30, Paragraphs("Valid"));
        PublisherDocument control = PublisherDocument.Create(300, 400);
        PublisherTextFrame expected = control.AddTextFrame(control.Pages[0].Id, 20, 20, 100, 30, Paragraphs("Valid"));
        Assert.Equal(expected.Id, created.Id); Assert.Equal(expected.StoryId, created.StoryId);
    }

    [Fact]
    public void Page_and_object_admission_limits_are_copied_and_do_not_count_pages_as_shape_objects() {
        var options = new PublisherReadOptions { MaximumPages = 2, Limits = new OfficeLegacyImportLimits { MaxItems = 2 } };
        PublisherDocument document = PublisherDocument.Create(300, 400, options);
        options.MaximumPages = 100; options.Limits.MaxItems = 100;
        PublisherPage master = document.AddMasterPage("Header");
        var pages = document.Pages; var masters = document.MasterPages;
        Assert.Throws<InvalidDataException>(() => document.AddPage());
        Assert.Throws<InvalidDataException>(() => document.AddMasterPage("Other"));
        Assert.Same(pages, document.Pages); Assert.Same(masters, document.MasterPages);
        PublisherTextFrame first = document.AddTextFrame(document.Pages[0].Id, 0, 0, 100, 20, Array.Empty<OfficeRichTextParagraph>());
        PublisherTextFrame second = document.AddTextFrame(master.Id, 0, 0, 100, 20, Array.Empty<OfficeRichTextParagraph>());
        Assert.Equal(2, document.ReadReport.CurrentObjectCount); Assert.Equal(2, document.ReadReport.TextStoryCount);
        var report = document.ReadReport;
        Assert.Throws<InvalidDataException>(() => document.AddTextFrame(document.Pages[0].Id, 0, 0, 100, 20, Array.Empty<OfficeRichTextParagraph>()));
        Assert.Same(report, document.ReadReport); Assert.Equal(first.StoryId + 1, second.Id);
    }

    [Fact]
    public void Repeated_master_projection_failure_does_not_publish_a_page_or_consume_its_identifier() {
        var options = new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 20 } };
        PublisherDocument document = PublisherDocument.Create(300, 400, options);
        PublisherPage master = document.AddMasterPage("Header");
        PublisherTextFrame header = document.AddTextFrame(master.Id, 12, 12, 200, 30, Paragraphs("Master"));
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        var pages = document.Pages; var report = document.ReadReport; string before = document.ToSvg();
        var failure = Assert.Throws<InvalidDataException>(() => document.AddPage(master.Id));
        Assert.Contains("projected text character limit", failure.Message);
        Assert.Same(pages, document.Pages); Assert.Same(report, document.ReadReport); Assert.Equal(before, document.ToSvg());
        PublisherPage blank = document.AddPage();
        Assert.Equal(header.StoryId + 1, blank.Id); Assert.Null(blank.MasterPageId);
        Assert.DoesNotContain("Master", document.ToSvg(1));
    }

    [Fact]
    public void Story_labels_count_toward_global_admission_and_failed_creation_can_be_retried() {
        var options = new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 12 } };
        PublisherDocument document = PublisherDocument.Create(300, 400, options);
        uint page = document.Pages[0].Id;
        // These frames deliberately retain overflow without painting characters,
        // so the assertion reaches semantic admission rather than scene limits.
        PublisherTextFrame first = document.AddTextFrame(page, 0, 0, .01, .01, Paragraphs("12345678"));
        var label = OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun("LABEL", 11, OfficeColor.Black), 0);
        var input = new[] { new OfficeRichTextParagraph(Array.Empty<OfficeRichTextRun>(), label) };
        var pages = document.Pages;
        Assert.Throws<ArgumentException>(() => document.AddTextFrame(page, 0, 0, .01, .01, input));
        Assert.Same(pages, document.Pages); Assert.Single(document.TextStories);
        document.ReplaceStoryParagraphs(first.StoryId, Array.Empty<OfficeRichTextParagraph>());
        PublisherTextFrame second = document.AddTextFrame(page, 0, 0, .01, .01, input);
        Assert.Equal(first.StoryId + 1, second.Id); Assert.Empty(document.TextStories.Single(story => story.Id == second.StoryId).Text);
        Assert.True(second.HasOverflow);
    }

    [Fact]
    public void Insufficient_pair_identity_rejects_a_frame_without_consuming_the_last_page_identifier() {
        PublisherDocument document = PublisherTableEditingFixture.Load("one-id");
        uint pageId = document.Pages[0].Id;
        var pages = document.Pages; var stories = document.TextStories;
        Assert.Throws<InvalidOperationException>(() => document.AddTextFrame(pageId, 12, 12, 100, 30, Paragraphs("Needs two IDs")));
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories);
        PublisherPage last = document.AddPage();
        Assert.Equal(uint.MaxValue, last.Id);
        Assert.Throws<InvalidOperationException>(() => document.AddMasterPage("No identifier"));
    }

    internal static OfficeRichTextParagraph[] Paragraphs(string text, bool bold = false) => new[] {
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, 11, OfficeColor.Black, bold: bold, fontFamily: "Arial") })
    };
    internal static string PaintedText(PublisherPage page) => string.Join("\n", PublisherNativeTests.Elements(page.CreateDrawing())
        .OfType<OfficeDrawingRichText>().Select(PublisherNativeTests.Text));
}
