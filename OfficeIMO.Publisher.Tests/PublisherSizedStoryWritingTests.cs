using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherSizedStoryWritingTests {
    private const string Body = "Neighbors share practical ideas, local news and plans for our shared spaces. Join the next meeting to help choose activities for the coming season.";
    private const string Columns = "Share your ideas. Meet your neighbors. Plan the next community project. Share your ideas. Meet your neighbors. Plan the next community project.";

    [Theory]
    [InlineData("native-edited.pub")]
    [InlineData("native-control-edited.pub")]
    public void Native_edited_portrait_linked_story_restores_through_public_save_without_changing_other_content(string fixture) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/GeneratedSizedComposition/" + fixture));
        PublisherDocument original = PublisherDocument.Load(source), document = PublisherDocument.Load(source);
        PublisherTextStory before = QualifiedBody(original, 'M');
        Assert.Equal(source, document.ToPubBytes());

        PublisherTextStory selected = document.TextStories.Single(story => story.Id == before.Id);
        OfficeRichTextParagraph paragraph = Assert.Single(selected.Paragraphs);
        OfficeRichTextRun[] runs = paragraph.Runs.ToArray();
        int first = Array.FindIndex(runs, run => run.Text.Length != 0);
        Assert.InRange(first, 0, runs.Length - 1);
        OfficeRichTextRun run = runs[first];
        Assert.StartsWith("M", run.Text);
        Assert.Null(run.LinkUri); Assert.Null(run.LinkDestinationName);
        runs[first] = new OfficeRichTextRun("N" + run.Text.Substring(1), run.FontSize, run.Color, run.Bold,
            run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle,
            run.StrikethroughStyle, run.Baseline, run.ParagraphIndent).WithLanguageId(run.LanguageId)
            .WithAutomaticPairKerningThreshold(run.AutomaticPairKerningThreshold);
        var replacement = new OfficeRichTextParagraph(runs, paragraph.Alignment, paragraph.LineHeight, paragraph.Margins,
            paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops);
        Assert.True(paragraph.SuppressAutomaticHyphenation.HasValue);
        replacement = replacement.WithAutomaticHyphenationSuppressed(paragraph.SuppressAutomaticHyphenation!.Value);
        document.ReplaceStoryParagraphs(selected.Id, new[] { replacement });

        byte[] saved = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument restored = PublisherDocument.Load(saved);
        Assert.Equal(before.Id, QualifiedBody(restored, 'N').Id);
        CompareStories(original, restored, before.Id);
        ComparePages(original.MasterPages.Concat(original.Pages).ToArray(), restored.MasterPages.Concat(restored.Pages).ToArray());
        Assert.Equal(original.Images.Select(image => (image.Id, image.ContentType)), restored.Images.Select(image => (image.Id, image.ContentType)));
        foreach (PublisherImage image in original.Images)
            Assert.Equal(image.GetBytes(), restored.Images.Single(item => item.Id == image.Id).GetBytes());
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved, "Quill/QuillSub/CONTENTS");
        Assert.Equal(saved, restored.ToPubBytes());
    }

    private static PublisherTextStory QualifiedBody(PublisherDocument document, char firstCharacter) {
        Assert.True(document.HasNativeSource);
        Assert.Equal(2, document.Pages.Count);
        PublisherPage master = Assert.Single(document.MasterPages);
        Assert.Equal(PublisherMasterPageSide.Single, master.MasterSide);
        Assert.Null(master.MasterPageId);
        PublisherPage[] owners = new[] { master }.Concat(document.Pages).ToArray();
        Assert.Equal(new[] { 3, 4, 3 }, owners.Select(page => page.ObjectIds.Count));
        Assert.Equal(10, owners.SelectMany(page => page.ObjectIds).Distinct().Count());
        Assert.All(owners, page => Assert.Equal((288D, 432D), (page.Width, page.Height)));
        Assert.All(document.Pages, page => { Assert.Equal(master.Id, page.MasterPageId); Assert.Null(page.MasterSide); });
        PublisherImage image = Assert.Single(document.Images);
        Assert.Equal("image/png", image.ContentType);
        PublisherPicture[] pictures = owners.SelectMany(page => page.Pictures).ToArray();
        Assert.Equal(2, pictures.Length);
        Assert.All(pictures, picture => { Assert.True(picture.HasPlacement); Assert.Equal(image.Id, picture.ImageId); });
        AssertBounds((234, 390, 36, 24), Assert.Single(master.Pictures));
        AssertBounds((216, 132, 54, 36), Assert.Single(document.Pages[0].Pictures));
        Assert.Empty(document.Pages[1].Pictures);

        string text = firstCharacter + Body.Substring(1) + "\n";
        PublisherTextStory body = Assert.Single(document.TextStories, story => story.Text == text);
        Assert.Equal(148, body.Text.Length);
        Assert.True(body.HasFinalParagraphTerminator);
        Assert.True(body.AutomaticHyphenation);
        AssertTypography(Assert.Single(body.Paragraphs), 10, 14, false);
        Assert.Equal(new[] { "Community update\n", "Events and projects\n", text, "Activity\nDate\nMeeting\nSaturday\n", Columns + "\n" }.OrderBy(value => value, StringComparer.Ordinal),
            document.TextStories.Where(story => story.Text.Any(character => !char.IsWhiteSpace(character))).Select(story => story.Text).OrderBy(value => value, StringComparer.Ordinal));

        PublisherTextFrame[] linked = document.Pages.SelectMany(page => page.TextFrames).Where(frame => frame.StoryId == body.Id).OrderBy(frame => frame.Order).ToArray();
        Assert.Equal(2, linked.Length);
        Assert.Equal(new uint[] { 0, 1 }, linked.Select(frame => frame.Order));
        Assert.Equal(document.Pages.Select(page => page.Id), linked.Select(frame => frame.PageId));
        Assert.Null(linked[0].PreviousFrameId); Assert.Null(linked[1].NextFrameId);
        Assert.Equal(linked[1].Id, linked[0].NextFrameId); Assert.Equal(linked[0].Id, linked[1].PreviousFrameId);
        AssertBounds((18, 90, 252, 14), linked[0]);
        AssertBounds((18, 48, 252, 56), linked[1]);
        Assert.All(linked, frame => {
            Assert.Equal((1, 0D, OfficeTextPadding.Empty, OfficeTextVerticalAlignment.Top),
                (frame.ColumnCount, frame.ColumnSpacing, frame.Padding, frame.VerticalAlignment));
            Assert.Empty(frame.WrapObjectIds);
        });
        AssertFrame(document, master, "Community update\n", (18, 12, 252, 14), 8, 12, false);
        AssertFrame(document, document.Pages[0], "Events and projects\n", (18, 48, 252, 30), 16, 20, true);
        PublisherTextFrame columns = AssertFrame(document, document.Pages[1], Columns + "\n", (18, 120, 252, 48), 9, 12, false);
        Assert.Equal((2, 12D, new OfficeTextPadding(3, 3, 3, 3), OfficeTextVerticalAlignment.Top),
            (columns.ColumnCount, columns.ColumnSpacing, columns.Padding, columns.VerticalAlignment));
        AssertRectangle(master, OfficeColor.Blue, (18, 30, 252, 3));
        AssertRectangle(document.Pages[1], OfficeColor.Red, (234, 378, 36, 12));

        PublisherTable table = Assert.Single(owners.SelectMany(page => page.Tables));
        Assert.Equal(document.Pages[0].Id, table.PageId);
        Assert.Equal((18D, 336D, 252D, 36D), (table.X, table.Y, table.Width, table.Height));
        Assert.Equal(new[] { 84D, 168D }, table.ColumnWidths);
        Assert.Equal(new[] { 18D, 18D }, table.RowHeights);
        Assert.Empty(table.BorderSegments);
        Assert.True(table.HasTextMapping);
        Assert.Equal(new[] { "Activity", "Date", "Meeting", "Saturday" }, table.Cells.Select(cell => cell.Text));
        for (int index = 0; index < table.Cells.Count; index++) {
            PublisherTableCell cell = table.Cells[index];
            Assert.Equal((index / 2, index % 2, 1, 1), (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan));
            AssertTypography(Assert.Single(cell.Paragraphs), 9, 12, index < 2);
        }
        return body;
    }

    private static PublisherTextFrame AssertFrame(PublisherDocument document, PublisherPage owner, string text,
        (double X, double Y, double Width, double Height) bounds, double fontSize, double lineHeight, bool bold) {
        PublisherTextStory story = Assert.Single(document.TextStories, item => item.Text == text);
        PublisherTextFrame frame = Assert.Single(owner.TextFrames, item => item.StoryId == story.Id);
        Assert.Equal(owner.Id, frame.PageId);
        Assert.Null(frame.PreviousFrameId); Assert.Null(frame.NextFrameId);
        Assert.Empty(frame.WrapObjectIds);
        AssertBounds(bounds, frame);
        AssertTypography(Assert.Single(story.Paragraphs), fontSize, lineHeight, bold);
        return frame;
    }

    private static void AssertTypography(OfficeRichTextParagraph paragraph, double fontSize, double lineHeight, bool bold) {
        Assert.Equal(OfficeTextAlignment.Left, paragraph.Alignment);
        Assert.Equal(lineHeight, paragraph.LineHeight);
        Assert.Null(paragraph.LineHeightFactor); Assert.Null(paragraph.Label);
        Assert.Equal(OfficeTextPadding.Empty, paragraph.Margins);
        Assert.Equal(OfficeTextParagraphIndent.Empty, paragraph.Indent);
        Assert.False(paragraph.SuppressAutomaticHyphenation);
        Assert.NotEmpty(paragraph.Runs);
        Assert.All(paragraph.Runs, run => {
            Assert.Equal((fontSize, "Arial", OfficeColor.Black, bold, false, OfficeTextDecorationStyle.None,
                OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal, (uint?)1033),
                (run.FontSize, run.FontFamily, run.Color, run.Bold, run.Italic, run.UnderlineStyle,
                    run.StrikethroughStyle, run.Baseline, run.LanguageId));
            Assert.Null(run.BackgroundColor); Assert.Null(run.LinkUri); Assert.Null(run.LinkDestinationName);
        });
    }

    private static void CompareStories(PublisherDocument before, PublisherDocument after, uint targetId) {
        Assert.Equal(before.TextStories.Select(story => story.Id), after.TextStories.Select(story => story.Id));
        foreach (PublisherTextStory original in before.TextStories) {
            PublisherTextStory restored = after.TextStories.Single(story => story.Id == original.Id);
            bool target = original.Id == targetId;
            Assert.Equal(target ? "N" + original.Text.Substring(1) : original.Text, restored.Text);
            Assert.Equal((original.AutomaticHyphenation, original.HasFinalParagraphTerminator),
                (restored.AutomaticHyphenation, restored.HasFinalParagraphTerminator));
            CompareParagraphs(original.Paragraphs, restored.Paragraphs, target);
        }
    }

    private static void CompareParagraphs(IReadOnlyList<OfficeRichTextParagraph> before, IReadOnlyList<OfficeRichTextParagraph> after, bool restoreFirstCharacter = false) {
        Assert.Equal(before.Count, after.Count);
        bool pending = restoreFirstCharacter;
        for (int index = 0; index < before.Count; index++) {
            OfficeRichTextParagraph a = before[index], b = after[index];
            Assert.Equal((a.Alignment, a.LineHeight, a.LineHeightFactor, a.Margins, a.Indent, a.SuppressAutomaticHyphenation),
                (b.Alignment, b.LineHeight, b.LineHeightFactor, b.Margins, b.Indent, b.SuppressAutomaticHyphenation));
            Assert.Null(a.Label); Assert.Null(b.Label);
            if (a.TabStops == null) Assert.Null(b.TabStops);
            else {
                OfficeTextTabStops tabs = Assert.IsType<OfficeTextTabStops>(b.TabStops);
                Assert.Equal((a.TabStops.DefaultInterval, a.TabStops.Origin, a.TabStops.AlignWithParagraph),
                    (tabs.DefaultInterval, tabs.Origin, tabs.AlignWithParagraph));
                Assert.Empty(a.TabStops.Stops); Assert.Empty(tabs.Stops);
            }
            Assert.Equal(a.Runs.Select(run => run.Text.Length), b.Runs.Select(run => run.Text.Length));
            for (int runIndex = 0; runIndex < a.Runs.Count; runIndex++) {
                OfficeRichTextRun x = a.Runs[runIndex], y = b.Runs[runIndex];
                string text = x.Text;
                if (pending && text.Length != 0) { text = "N" + text.Substring(1); pending = false; }
                Assert.Equal(text, y.Text);
                // Includes empty typing carriers and the final run supplying CR styling.
                Assert.Equal(Style(x), Style(y));
            }
        }
        Assert.False(pending);
    }

    private static object Style(OfficeRichTextRun run) => new {
        run.FontSize, run.Color, run.Bold, run.Italic, run.FontFamily, run.UnderlineStyle,
        run.StrikethroughStyle, run.BackgroundColor, run.Baseline, run.ParagraphIndent,
        run.LanguageId, run.AutomaticPairKerningThreshold, run.LinkUri, run.LinkDestinationName
    };

    private static void ComparePages(PublisherPage[] before, PublisherPage[] after) {
        Assert.Equal(before.Length, after.Length);
        for (int index = 0; index < before.Length; index++) {
            PublisherPage a = before[index], b = after[index];
            Assert.Equal((a.Id, a.Name, a.MasterPageId, a.MasterSide, a.Width, a.Height, a.HasNativeOrigin),
                (b.Id, b.Name, b.MasterPageId, b.MasterSide, b.Width, b.Height, b.HasNativeOrigin));
            Assert.Equal(a.ObjectIds, b.ObjectIds);
            Assert.Equal(a.TextFrames.Select(frame => frame.Id), b.TextFrames.Select(frame => frame.Id));
            foreach (PublisherTextFrame frame in a.TextFrames) CompareFrame(frame, b.TextFrames.Single(item => item.Id == frame.Id));
            Assert.Equal(a.Tables.Select(table => table.Id), b.Tables.Select(table => table.Id));
            foreach (PublisherTable table in a.Tables) CompareTable(table, b.Tables.Single(item => item.Id == table.Id));
            Assert.Equal(a.Pictures.Select(picture => picture.Id), b.Pictures.Select(picture => picture.Id));
            foreach (PublisherPicture picture in a.Pictures) {
                PublisherPicture actual = b.Pictures.Single(item => item.Id == picture.Id);
                Assert.Equal((picture.PageId, picture.ImageId, picture.X, picture.Y, picture.Width, picture.Height, picture.PageTransform),
                    (actual.PageId, actual.ImageId, actual.X, actual.Y, actual.Width, actual.Height, actual.PageTransform));
                Assert.Equal((picture.Crop.Left, picture.Crop.Top, picture.Crop.Right, picture.Crop.Bottom, picture.HasPlacement, picture.CanEdit, picture.EditLimitation),
                    (actual.Crop.Left, actual.Crop.Top, actual.Crop.Right, actual.Crop.Bottom, actual.HasPlacement, actual.CanEdit, actual.EditLimitation));
            }
            Assert.Equal(a.Shapes.Select(shape => shape.Id), b.Shapes.Select(shape => shape.Id));
            foreach (PublisherShape shape in a.Shapes) {
                PublisherShape actual = b.Shapes.Single(item => item.Id == shape.Id);
                Assert.Equal((shape.PageId, shape.X, shape.Y, shape.Width, shape.Height, shape.PageTransform, shape.HasPlacement, shape.CanEdit, shape.EditLimitation),
                    (actual.PageId, actual.X, actual.Y, actual.Width, actual.Height, actual.PageTransform, actual.HasPlacement, actual.CanEdit, actual.EditLimitation));
                if (shape.CanEdit) ComparePaint(shape.CreateShape(), actual.CreateShape());
            }
        }
    }

    private static void CompareFrame(PublisherTextFrame a, PublisherTextFrame b) {
        Assert.Equal((a.PageId, a.StoryId, a.PreviousFrameId, a.NextFrameId, a.Order),
            (b.PageId, b.StoryId, b.PreviousFrameId, b.NextFrameId, b.Order));
        Assert.Equal((a.X, a.Y, a.Width, a.Height, a.PageTransform), (b.X, b.Y, b.Width, b.Height, b.PageTransform));
        Assert.Equal((a.ColumnCount, a.ColumnSpacing, a.Padding, a.VerticalAlignment, a.CanEditLayout, a.LayoutEditLimitation),
            (b.ColumnCount, b.ColumnSpacing, b.Padding, b.VerticalAlignment, b.CanEditLayout, b.LayoutEditLimitation));
        Assert.Equal((a.TextStart, a.TextLength, a.HasOverflow), (b.TextStart, b.TextLength, b.HasOverflow));
        Assert.Equal(a.WrapObjectIds, b.WrapObjectIds);
    }

    private static void CompareTable(PublisherTable a, PublisherTable b) {
        Assert.Equal((a.PageId, a.StoryId, a.X, a.Y, a.Width, a.Height, a.PageTransform),
            (b.PageId, b.StoryId, b.X, b.Y, b.Width, b.Height, b.PageTransform));
        Assert.Equal((a.HasPlacement, a.HasTextMapping, a.CanEditCellText, a.CellTextEditLimitation, a.CanEditBounds, a.BoundsEditLimitation),
            (b.HasPlacement, b.HasTextMapping, b.CanEditCellText, b.CellTextEditLimitation, b.CanEditBounds, b.BoundsEditLimitation));
        Assert.Equal(a.ColumnWidths, b.ColumnWidths); Assert.Equal(a.RowHeights, b.RowHeights);
        Assert.Equal(a.BorderSegments, b.BorderSegments);
        Assert.Equal(a.Cells.Count, b.Cells.Count);
        for (int index = 0; index < a.Cells.Count; index++) {
            PublisherTableCell x = a.Cells[index], y = b.Cells[index];
            Assert.Equal((x.RowIndex, x.ColumnIndex, x.RowSpan, x.ColumnSpan, x.X, x.Y, x.Width, x.Height, x.FillColor, x.Text),
                (y.RowIndex, y.ColumnIndex, y.RowSpan, y.ColumnSpan, y.X, y.Y, y.Width, y.Height, y.FillColor, y.Text));
            CompareParagraphs(x.Paragraphs, y.Paragraphs);
        }
    }

    private static void ComparePaint(OfficeShape a, OfficeShape b) {
        Assert.Equal((a.Kind, a.Width, a.Height, a.FillColor, a.StrokeColor, a.StrokeWidth, a.FillOpacity, a.StrokeOpacity),
            (b.Kind, b.Width, b.Height, b.FillColor, b.StrokeColor, b.StrokeWidth, b.FillOpacity, b.StrokeOpacity));
        Assert.Equal((a.StrokeDashStyle, a.StrokeDashOffset, a.StrokeMiterLimit, a.StrokeLineCap, a.StrokeLineJoin, a.FillRule, a.CornerRadius, a.Transform),
            (b.StrokeDashStyle, b.StrokeDashOffset, b.StrokeMiterLimit, b.StrokeLineCap, b.StrokeLineJoin, b.FillRule, b.CornerRadius, b.Transform));
        Assert.Equal(a.StrokeDashArray, b.StrokeDashArray); Assert.Equal(a.Points, b.Points); Assert.Equal(a.PathCommands, b.PathCommands);
        Assert.Null(a.FillGradient); Assert.Null(b.FillGradient); Assert.Null(a.FillRadialGradient); Assert.Null(b.FillRadialGradient);
        Assert.Null(a.StrokeGradient); Assert.Null(b.StrokeGradient); Assert.Null(a.StrokeRadialGradient); Assert.Null(b.StrokeRadialGradient);
        Assert.Null(a.StrokeStartMarker); Assert.Null(b.StrokeStartMarker); Assert.Null(a.StrokeEndMarker); Assert.Null(b.StrokeEndMarker);
        Assert.Null(a.Shadow); Assert.Null(b.Shadow); Assert.Null(a.Glow); Assert.Null(b.Glow); Assert.Null(a.ClipPath); Assert.Null(b.ClipPath);
    }

    private static void AssertRectangle(PublisherPage page, OfficeColor fill, (double X, double Y, double Width, double Height) bounds) {
        PublisherShape shape = Assert.Single(page.Shapes, item => item.CanEdit && item.CreateShape().FillColor == fill);
        OfficeShape paint = shape.CreateShape();
        Assert.Equal(OfficeShapeKind.Rectangle, paint.Kind); Assert.Null(paint.StrokeColor);
        AssertBounds(bounds, (shape.X, shape.Y, shape.Width, shape.Height));
    }

    private static void AssertBounds((double X, double Y, double Width, double Height) expected, PublisherTextFrame actual) =>
        AssertBounds(expected, (actual.X, actual.Y, actual.Width, actual.Height));
    private static void AssertBounds((double X, double Y, double Width, double Height) expected, PublisherPicture actual) =>
        AssertBounds(expected, (actual.X, actual.Y, actual.Width, actual.Height));
    private static void AssertBounds((double X, double Y, double Width, double Height) expected, (double X, double Y, double Width, double Height) actual) {
        Assert.InRange(actual.X, expected.X - 0.001, expected.X + 0.001);
        Assert.InRange(actual.Y, expected.Y - 0.001, expected.Y + 0.001);
        Assert.InRange(actual.Width, expected.Width - 0.001, expected.Width + 0.001);
        Assert.InRange(actual.Height, expected.Height - 0.001, expected.Height + 0.001);
    }
}
