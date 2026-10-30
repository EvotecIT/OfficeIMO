using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherFrameLayoutTests {
    [Fact]
    public void Authored_columns_reflow_one_complete_styled_story_and_continuation_inherits_current_layout() {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        string body = string.Concat(Enumerable.Repeat("Column content remains complete. ", 70));
        var paragraphs = new[] { new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("Column heading\n", 14, OfficeColor.Black, bold: true, fontFamily: "Arial"),
            new OfficeRichTextRun(body, 11, OfficeColor.Blue, italic: true, fontFamily: "Arial")
        }) };
        PublisherTextFrame original = document.AddTextFrame(document.Pages[0].Id, 36, 72, 360, 180, paragraphs);
        PublisherTextStory story = Assert.Single(document.TextStories);
        PublisherPage oldPage = document.Pages[0]; string oldSvg = document.ToSvg(); var source = document.SourceReadReport;
        var padding = new OfficeTextPadding(8, 6, 10, 12);
        document.SetTextFrameLayout(original.Id, 2, 18, padding, OfficeTextVerticalAlignment.Top);
        PublisherTextFrame edited = Assert.Single(document.Pages[0].TextFrames);
        Assert.True(edited.CanEditLayout, edited.LayoutEditLimitation); Assert.Null(edited.LayoutEditLimitation);
        Assert.Equal(2, edited.ColumnCount); Assert.Equal(18, edited.ColumnSpacing); Assert.Equal(padding, edited.Padding);
        Assert.Equal(OfficeTextVerticalAlignment.Top, edited.VerticalAlignment); Assert.Equal(Placement(original), Placement(edited));
        OfficeDrawingRichText[] columns = Painted(document.Pages[0], original.Id);
        Assert.Equal(2, columns.Length);
        Assert.Equal(44, columns[0].X, 6); Assert.Equal(78, columns[0].Y, 6);
        Assert.Equal(162, columns[0].Width, 6); Assert.Equal(162, columns[0].Height, 6);
        Assert.Equal(224, columns[1].X, 6); Assert.Equal(78, columns[1].Y, 6);
        Assert.Equal(columns[0].Width, columns[1].Width, 6); Assert.True(edited.HasOverflow);
        Assert.True(columns[0].Paragraphs[0].Runs[0].Bold);
        Assert.Contains(columns.SelectMany(column => column.Paragraphs).SelectMany(paragraph => paragraph.Runs), run => run.Italic && run.Color == OfficeColor.Blue);
        PublisherPage destination = document.AddPage();
        PublisherTextFrame continuation = document.AddLinkedTextFrame(original.Id, destination.Id, 36, 72, 540, 640);
        Assert.Equal(2, continuation.ColumnCount); Assert.Equal(18, continuation.ColumnSpacing); Assert.Equal(padding, continuation.Padding);
        Assert.Equal(OfficeTextVerticalAlignment.Top, continuation.VerticalAlignment);
        Assert.Equal(story.Text, Assigned(document, story.Id)); Assert.False(continuation.HasOverflow);
        Assert.Same(story, Assert.Single(document.TextStories)); Assert.Same(source, document.SourceReadReport);
        Assert.Equal(1, original.ColumnCount); Assert.Equal(OfficeTextPadding.Empty, original.Padding);
        Assert.Equal(oldSvg, OfficeDrawingSvgExporter.ToSvg(oldPage.CreateDrawing()));
        document.SetTextFrameLayout(original.Id, 1, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Center);
        PublisherTextFrame tail = Assert.Single(document.Pages[1].TextFrames);
        Assert.Equal(2, tail.ColumnCount); Assert.Equal(padding, tail.Padding); Assert.Equal(OfficeTextVerticalAlignment.Top, tail.VerticalAlignment);
        Assert.Equal(story.Text, Assigned(document, story.Id));
        document.ReplaceStoryParagraphs(story.Id, PublisherAuthoringTests.Paragraphs("Short"));
        Assert.Equal(0, Assert.Single(document.Pages[1].TextFrames).TextLength);
        Assert.Equal("Short", Assert.Single(document.TextStories).Text); Assert.Equal("Column heading\n" + body, story.Text);
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Equal(2, pdf.Pages.Count); Assert.Contains("Short", pdf.Pages[0].ExtractText());
    }

    [Fact]
    public void Loaded_layout_and_master_alignment_reach_exports_without_changing_roles_or_frame_links() {
        PublisherDocument document = Load("MultipleMasters");
        PublisherTextFrame physical = document.Pages[0].TextFrames[0], master = document.MasterPages[0].TextFrames[0];
        var source = document.SourceReadReport; int count = document.ReadReport.CurrentObjectCount;
        uint?[] assignments = document.Pages.Select(page => page.MasterPageId).ToArray();
        PublisherTextStory[] stories = document.TextStories.ToArray();
        var padding = new OfficeTextPadding(6, 4, 8, 6);
        document.SetTextFrameLayout(physical.Id, 2, 12, padding, OfficeTextVerticalAlignment.Top);
        document.SetTextFrameLayout(master.Id, 1, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Bottom);
        PublisherTextFrame now = document.Pages[0].TextFrames[0], currentMaster = document.MasterPages[0].TextFrames[0];
        Assert.Equal(Placement(physical), Placement(now)); Assert.Equal(Placement(master), Placement(currentMaster));
        Assert.Equal(assignments, document.Pages.Select(page => page.MasterPageId)); Assert.Equal(stories, document.TextStories);
        Assert.Equal(count, document.ReadReport.CurrentObjectCount); Assert.Same(source, document.SourceReadReport);
        Assert.All(document.Pages, page => Assert.DoesNotContain(page.TextFrames, frame => frame.Id == master.Id));
        Assert.All(new[] { 0, 2 }, index => {
            OfficeDrawingRichText heading = Assert.Single(Painted(document.Pages[index], master.Id));
            Assert.Equal(OfficeTextVerticalAlignment.Bottom, heading.VerticalAlignment);
            Assert.Contains("Master Alpha", document.ToSvg(index));
        });
        Assert.Empty(Painted(document.Pages[1], master.Id));
        Assert.Equal(OfficeTextVerticalAlignment.Top, master.VerticalAlignment);
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Contains("Master Alpha", pdf.Pages[0].ExtractText()); Assert.Contains("Master Alpha", pdf.Pages[2].ExtractText());
        Assert.DoesNotContain("Master Alpha", pdf.Pages[1].ExtractText());
    }

    [Fact]
    public void Vertical_alignment_moves_the_same_short_paragraph_inside_unchanged_insets_and_bounds() {
        PublisherDocument document = PublisherDocument.Create(300, 400);
        PublisherTextFrame frame = document.AddTextFrame(document.Pages[0].Id, 20, 30, 240, 120, PublisherAuthoringTests.Paragraphs("Alignment marker"));
        var padding = new OfficeTextPadding(6, 8, 10, 12);
        var baselines = new List<double>();
        foreach (OfficeTextVerticalAlignment alignment in new[] { OfficeTextVerticalAlignment.Top, OfficeTextVerticalAlignment.Center, OfficeTextVerticalAlignment.Bottom }) {
            document.SetTextFrameLayout(frame.Id, 1, 0, padding, alignment);
            PublisherTextFrame current = Assert.Single(document.Pages[0].TextFrames);
            Assert.Equal(Placement(frame), Placement(current)); Assert.Equal(padding, current.Padding); Assert.Equal(alignment, current.VerticalAlignment);
            OfficeDrawingRichText paint = Assert.Single(Painted(document.Pages[0], frame.Id));
            Assert.Equal(26, paint.X); Assert.Equal(38, paint.Y); Assert.Equal(224, paint.Width); Assert.Equal(100, paint.Height);
            Assert.Equal(alignment, paint.VerticalAlignment);
            PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
            PdfTextSpan span = Assert.Single(pdf.Pages[0].GetTextSpans(), item => item.IsVisible && item.Text == "Alignment marker");
            Assert.InRange(span.X, 26, 26.1); baselines.Add(span.Y);
        }
        Assert.True(baselines[0] > baselines[1]); Assert.True(baselines[1] > baselines[2]);
        Assert.Equal(baselines[0] - baselines[1], baselines[1] - baselines[2], 4);
    }

    [Fact]
    public void Mixed_text_artwork_keeps_one_placement_and_layout_correction_can_recover_a_collapsed_content_area() {
        PublisherDocument document = PublisherFrameCreationFixture.Load(mixedPicture: true).Document;
        PublisherTextFrame frame = Assert.Single(document.Pages[0].TextFrames); PublisherPicture picture = Assert.Single(document.Pages[0].Pictures);
        Assert.True(frame.CanEditLayout, frame.LayoutEditLimitation); Assert.False(picture.CanEdit);
        document.SetTextFrameLayout(frame.Id, 2, 12, new OfficeTextPadding(8, 4, 8, 4), OfficeTextVerticalAlignment.Center);
        PublisherTextFrame current = Assert.Single(document.Pages[0].TextFrames); PublisherPicture currentPicture = Assert.Single(document.Pages[0].Pictures);
        Assert.Equal(Placement(frame), Placement(current)); Assert.Equal((picture.X, picture.Y, picture.Width, picture.Height, picture.PageTransform),
            (currentPicture.X, currentPicture.Y, currentPicture.Width, currentPicture.Height, currentPicture.PageTransform));
        Assert.False(currentPicture.CanEdit); Assert.Equal(picture.ImageId, currentPicture.ImageId);
        PublisherDocument rectangle = Load("Rectangle");
        PublisherShape graphic = Assert.Single(rectangle.Pages[0].Shapes); PublisherTextFrame text = Assert.Single(rectangle.Pages[0].TextFrames);
        rectangle.SetTextFrameLayout(text.Id, 2, 6, new OfficeTextPadding(2, 2, 2, 2), OfficeTextVerticalAlignment.Bottom);
        Assert.Equal((graphic.X, graphic.Y, graphic.Width, graphic.Height, graphic.PageTransform),
            (rectangle.Pages[0].Shapes[0].X, rectangle.Pages[0].Shapes[0].Y, rectangle.Pages[0].Shapes[0].Width, rectangle.Pages[0].Shapes[0].Height, rectangle.Pages[0].Shapes[0].PageTransform));
        Assert.Equal(graphic.CreateShape().FillColor, rectangle.Pages[0].Shapes[0].CreateShape().FillColor);

        PublisherDocument collapsed = Load("Text"); PublisherTextFrame original = Assert.Single(collapsed.Pages[0].TextFrames);
        collapsed.SetTextFrameLayout(original.Id, 1, 0, new OfficeTextPadding(80, 4, 80, 4), OfficeTextVerticalAlignment.Top);
        collapsed.SetTextFrameBounds(original.Id, original.X, original.Y, 120, original.Height);
        PublisherTextFrame narrow = Assert.Single(collapsed.Pages[0].TextFrames);
        Assert.True(narrow.CanEditLayout, narrow.LayoutEditLimitation); Assert.True(narrow.HasOverflow);
        Assert.Contains(collapsed.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_FRAME_HAS_NO_CONTENT_AREA");
        collapsed.SetTextFrameLayout(original.Id, 1, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top);
        Assert.False(Assert.Single(collapsed.Pages[0].TextFrames).HasOverflow);
        Assert.Contains("Publisher text", collapsed.ToSvg());
        Assert.DoesNotContain(collapsed.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_FRAME_HAS_NO_CONTENT_AREA");
    }

    [Theory]
    [InlineData("anchor")]
    [InlineData("direction")]
    [InlineData("wrapping")]
    [InlineData("fitting")]
    public void Unassessed_imported_layout_overrides_keep_reading_and_bounds_but_explicitly_reject_layout_editing(string kind) {
        ushort property = kind == "anchor" ? (ushort)0x87 : kind == "direction" ? (ushort)0x88 : kind == "wrapping" ? (ushort)0x85 : (ushort)0xBF;
        uint value = kind == "anchor" || kind == "wrapping" ? 0U : kind == "direction" ? 1U : 0x00020002U;
        byte[] input = PublisherDrawingFixture.Mutate(new Dictionary<ushort, uint> { [property] = value }, 202,
            fixture: "NativeComparison/Text/native-text.pub");
        PublisherDocument document = PublisherDocument.Load(input); PublisherTextFrame frame = Assert.Single(document.Pages[0].TextFrames);
        Assert.False(frame.CanEditLayout); Assert.NotNull(frame.LayoutEditLimitation); Assert.Contains("Publisher text", document.ToSvg());
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_FRAME_LAYOUT_EDIT_UNSUPPORTED" && item.Message == frame.LayoutEditLimitation);
        var pages = document.Pages; var report = document.ReadReport;
        Assert.Throws<NotSupportedException>(() => document.SetTextFrameLayout(frame.Id, 2, 12, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Same(pages, document.Pages); Assert.Same(report, document.ReadReport);
        document.SetTextFrameBounds(frame.Id, frame.X + 6, frame.Y, frame.Width, frame.Height);
        Assert.Equal(frame.X + 6, Assert.Single(document.Pages[0].TextFrames).X, 6);
        Assert.False(Assert.Single(document.Pages[0].TextFrames).CanEditLayout);
    }

    [Fact]
    public void Tables_aliases_and_grouped_frames_reject_layout_without_changing_the_publication() {
        PublisherDocument alias = PublisherTableEditingFixture.Load("ordinary-frame");
        PublisherTextFrame frame = Assert.Single(alias.Pages.SelectMany(page => page.TextFrames), item => item.Id == 902);
        Assert.False(frame.CanEditLayout); Assert.Contains("table", frame.LayoutEditLimitation!);
        var pages = alias.Pages;
        Assert.Throws<NotSupportedException>(() => alias.SetTextFrameLayout(frame.Id, 1, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Throws<ArgumentOutOfRangeException>(() => alias.SetTextFrameLayout(alias.Pages.SelectMany(page => page.Tables).First().Id, 1, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Same(pages, alias.Pages);
        PublisherDocument grouped = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        PublisherTextFrame groupedFrame = grouped.Pages.SelectMany(page => page.TextFrames).Single(item => item.Id == 345);
        Assert.False(groupedFrame.CanEditLayout); var groupedPages = grouped.Pages;
        Assert.Throws<NotSupportedException>(() => grouped.SetTextFrameLayout(groupedFrame.Id, 1, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Same(groupedPages, grouped.Pages);
    }

    [Fact]
    public void Unknown_native_vertical_alignment_keeps_its_read_fallback_and_bounds_contract_but_rejects_layout_edits() {
        PublisherDocument baseline = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        PublisherTextFrame eligible = baseline.Pages.Concat(baseline.MasterPages).SelectMany(page => page.TextFrames)
            .First(frame => frame.CanEditLayout && frame.VerticalAlignment != OfficeTextVerticalAlignment.Top);
        // Existing corpus fields 0x35 use the four-byte scalar encoding.
        byte[] input = PublisherInputContractTests.Mutate("Contents", bytes =>
            PublisherInputContractTests.WriteUInt32(bytes, PublisherTextFlowTests.NativeFrameField(bytes, eligible.Id, 0x35), 99), "SampleNewsletter.pub");
        PublisherDocument document = PublisherDocument.Load(input);
        PublisherTextFrame frame = document.Pages.Concat(document.MasterPages).SelectMany(page => page.TextFrames).Single(item => item.Id == eligible.Id);
        Assert.Equal(OfficeTextVerticalAlignment.Top, frame.VerticalAlignment);
        Assert.False(frame.CanEditLayout); Assert.Contains("native vertical alignment", frame.LayoutEditLimitation!);
        var pages = document.Pages;
        Assert.Throws<NotSupportedException>(() => document.SetTextFrameLayout(frame.Id, 1, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Same(pages, document.Pages);
        document.SetTextFrameBounds(frame.Id, frame.X, frame.Y, frame.Width, frame.Height);
        Assert.False(document.Pages.Concat(document.MasterPages).SelectMany(page => page.TextFrames).Single(item => item.Id == frame.Id).CanEditLayout);
    }

    [Fact]
    public void Invalid_layout_arithmetic_cancellation_and_resource_failure_preserve_the_previous_model_and_scene() {
        PublisherDocument document = PublisherDocument.Create(300, 400);
        PublisherTextFrame frame = document.AddTextFrame(document.Pages[0].Id, 20, 20, 240, 100, PublisherAuthoringTests.Paragraphs("Unchanged"));
        var pages = document.Pages; var stories = document.TextStories; var report = document.ReadReport; string svg = document.ToSvg();
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetTextFrameLayout(frame.Id, 0, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetTextFrameLayout(frame.Id, 257, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetTextFrameLayout(frame.Id, 1, double.NaN, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetTextFrameLayout(frame.Id, 1, -1, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetTextFrameLayout(frame.Id, 1, 0, OfficeTextPadding.Empty, (OfficeTextVerticalAlignment)99));
        Assert.Throws<ArgumentException>(() => document.SetTextFrameLayout(frame.Id, 2, 240, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Throws<ArgumentException>(() => document.SetTextFrameLayout(frame.Id, 1, 0, new OfficeTextPadding(0, 50, 0, 50), OfficeTextVerticalAlignment.Top));
        Assert.Throws<ArgumentException>(() => document.SetTextFrameLayout(frame.Id, 1, 0, new OfficeTextPadding(double.MaxValue, 0, double.MaxValue, 0), OfficeTextVerticalAlignment.Top));
        Assert.Throws<ArgumentException>(() => document.SetTextFrameLayout(frame.Id, 256, double.MaxValue, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetTextFrameLayout(999, 1, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Throws<OperationCanceledException>(() => document.SetTextFrameLayout(frame.Id, 1, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top, new CancellationToken(true)));
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories); Assert.Same(report, document.ReadReport); Assert.Equal(svg, document.ToSvg());

        PublisherDocument bounded = PublisherDocument.Create(300, 400, new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 512 } });
        string body = string.Concat(Enumerable.Repeat("Bounded continuation text. ", 16));
        PublisherTextFrame narrow = bounded.AddTextFrame(bounded.Pages[0].Id, 20, 20, 40, 32, PublisherAuthoringTests.Paragraphs(body));
        var boundedPages = bounded.Pages; var boundedReport = bounded.ReadReport; string before = bounded.ToSvg();
        InvalidDataException failure = Assert.Throws<InvalidDataException>(() => bounded.SetTextFrameLayout(narrow.Id, 2, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Contains("text layout character work limit", failure.Message);
        Assert.Same(boundedPages, bounded.Pages); Assert.Same(boundedReport, bounded.ReadReport); Assert.Equal(before, bounded.ToSvg());
        bounded.SetTextFrameLayout(narrow.Id, 1, 0, new OfficeTextPadding(2, 2, 2, 2), OfficeTextVerticalAlignment.Top);
        Assert.Equal(1, Assert.Single(bounded.Pages[0].TextFrames).ColumnCount);

        PublisherDocument items = PublisherDocument.Create(300, 400, new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = 2 } });
        PublisherTextFrame empty = items.AddTextFrame(items.Pages[0].Id, 0, 0, 200, 100, Array.Empty<OfficeRichTextParagraph>());
        var itemPages = items.Pages;
        Assert.Throws<InvalidDataException>(() => items.SetTextFrameLayout(empty.Id, 3, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top));
        Assert.Same(itemPages, items.Pages);
        items.SetTextFrameLayout(empty.Id, 2, 0, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top);
        Assert.Equal(2, Assert.Single(items.Pages[0].TextFrames).ColumnCount);
    }

    private static PublisherDocument Load(string name) => PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/" + name + "/native-" + name.ToLowerInvariant() + ".pub"));
    private static OfficeDrawingRichText[] Painted(PublisherPage page, uint id) => PublisherNativeTests.Elements(page.CreateDrawing())
        .OfType<OfficeDrawingRichText>().Where(text => text.SourceElementIds?.Contains("publisher-object-" + id) == true).ToArray();
    private static (uint Id, uint Story, uint Page, uint? Previous, uint? Next, uint Order, double X, double Y, double Width, double Height, OfficeTransform Transform) Placement(PublisherTextFrame frame) =>
        (frame.Id, frame.StoryId, frame.PageId, frame.PreviousFrameId, frame.NextFrameId, frame.Order, frame.X, frame.Y, frame.Width, frame.Height, frame.PageTransform);
    private static string Assigned(PublisherDocument document, uint storyId) {
        PublisherTextStory story = document.TextStories.Single(item => item.Id == storyId);
        return string.Concat(document.Pages.SelectMany(page => page.TextFrames).Where(frame => frame.StoryId == storyId).OrderBy(frame => frame.Order)
            .Select(frame => story.Text.Substring(frame.TextStart!.Value, frame.TextLength!.Value)));
    }
}
