using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPairedCompositionTests {
    private const long MaximumBytes = 64L * 1024 * 1024;

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Composed_creation_keeps_master_registration_artwork_and_physical_story_table_ownership(bool paired) {
        PublisherDocument document = Author(paired);
        int[] assignments = { 0, 1, paired ? 2 : 0 };
        byte[] image = Assert.Single(document.Images).GetBytes();
        PublisherTextFrame first = document.Pages[0].TextFrames.Single(frame => frame.NextFrameId.HasValue);
        PublisherTextFrame columns = document.Pages[0].TextFrames.Single(frame => frame.ColumnCount == 2);

        byte[] generated = PublisherNativeCreationCodec.Create(PublisherNativeCreationProjection.ComposedContent(document),
            new PublisherReadOptions(), MaximumBytes, CancellationToken.None);
        PublisherDocument decoded = PublisherDocument.Load(generated);

        Assert.Equal(document.MasterPages.Select(page => (page.Name, page.MasterSide)),
            decoded.MasterPages.Select(page => (page.Name, page.MasterSide)));
        Assert.Equal(assignments, MasterIndices(decoded));
        Assert.Equal(document.Pages.Select(page => page.ObjectIds.Count), decoded.Pages.Select(page => page.ObjectIds.Count));
        Assert.Equal(document.TextStories.Count, decoded.TextStories.Count);
        PublisherImage asset = Assert.Single(decoded.Images);
        Assert.Equal("image/png", asset.ContentType);
        Assert.Equal(image, asset.GetBytes());
        Assert.Equal(document.MasterPages.Count, decoded.MasterPages.Concat(decoded.Pages).Sum(page => page.Pictures.Count));
        Assert.All(decoded.MasterPages.Concat(decoded.Pages).SelectMany(page => page.Pictures),
            picture => Assert.Equal(asset.Id, picture.ImageId));

        for (int index = 0; index < document.MasterPages.Count; index++) {
            PublisherPage expected = document.MasterPages[index], actual = decoded.MasterPages[index];
            Assert.Null(actual.MasterPageId);
            Assert.Equal((expected.Width, expected.Height), (actual.Width, actual.Height));
            Assert.Equal(expected.ObjectIds.Count, actual.ObjectIds.Count);
            Assert.Empty(actual.Tables);
            PublisherTextFrame frame = Assert.Single(actual.TextFrames);
            Assert.Equal(actual.Id, frame.PageId);
            Assert.Equal(Bounds(expected.TextFrames[0]), Bounds(frame));
            AssertStory(document.TextStories.Single(story => story.Id == expected.TextFrames[0].StoryId),
                decoded.TextStories.Single(story => story.Id == frame.StoryId));
            PublisherShape rectangle = Assert.Single(actual.Shapes, shape => shape.Id != frame.Id);
            PublisherShape expectedRectangle = Assert.Single(expected.Shapes);
            Assert.Equal(actual.Id, rectangle.PageId);
            Assert.Equal((expectedRectangle.X, expectedRectangle.Y, expectedRectangle.Width, expectedRectangle.Height),
                (rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height));
            Assert.Equal(OfficeShapeKind.Rectangle, rectangle.CreateShape().Kind);
            Assert.Equal(expectedRectangle.CreateShape().FillColor, rectangle.CreateShape().FillColor);
            Assert.Null(rectangle.CreateShape().StrokeColor);
            PublisherPicture picture = Assert.Single(actual.Pictures), expectedPicture = Assert.Single(expected.Pictures);
            Assert.Equal(actual.Id, picture.PageId);
            Assert.Equal((expectedPicture.X, expectedPicture.Y, expectedPicture.Width, expectedPicture.Height),
                (picture.X, picture.Y, picture.Width, picture.Height));
        }

        uint[] masterObjects = decoded.MasterPages.SelectMany(page => page.ObjectIds).ToArray();
        for (int index = 0; index < decoded.Pages.Count; index++) {
            PublisherPage page = decoded.Pages[index];
            Assert.Null(page.MasterSide);
            Assert.Equal((612D, 792D), (page.Width, page.Height));
            Assert.Empty(page.Pictures);
            Assert.All(page.TextFrames, frame => Assert.Equal(page.Id, frame.PageId));
            Assert.All(page.Shapes, shape => Assert.Equal(page.Id, shape.PageId));
            Assert.All(page.Tables, item => Assert.Equal(page.Id, item.PageId));
            Assert.DoesNotContain(page.ObjectIds, masterObjects.Contains);
        }
        uint[] owned = decoded.MasterPages.Concat(decoded.Pages).SelectMany(page => page.ObjectIds).ToArray();
        Assert.Equal(owned.Length, owned.Distinct().Count());

        PublisherTextFrame[] chain = decoded.Pages.SelectMany(page => page.TextFrames)
            .Where(frame => frame.PreviousFrameId.HasValue || frame.NextFrameId.HasValue).OrderBy(frame => frame.Order).ToArray();
        Assert.Equal(3, chain.Length);
        PublisherTextStory linked = decoded.TextStories.Single(story => story.Id == chain[0].StoryId);
        AssertStory(document.TextStories.Single(story => story.Id == first.StoryId), linked);
        Assert.All(chain, frame => Assert.Equal(linked.Id, frame.StoryId));
        Assert.Equal(new uint[] { 0, 1, 2 }, chain.Select(frame => frame.Order));
        var pageIndices = decoded.Pages.Select((page, index) => (page.Id, index)).ToDictionary(item => item.Id, item => item.index);
        Assert.Equal(new[] { 0, 2, 1 }, chain.Select(frame => pageIndices[frame.PageId]));
        Assert.Equal(document.Pages.SelectMany(page => page.TextFrames).Where(frame => frame.StoryId == first.StoryId)
            .OrderBy(frame => frame.Order).Select(Bounds), chain.Select(Bounds));
        Assert.Null(chain[0].PreviousFrameId); Assert.Null(chain[2].NextFrameId);
        Assert.Equal(0, chain[0].TextStart);
        Assert.All(chain, frame => { Assert.True(frame.TextStart.HasValue); Assert.True(frame.TextLength > 0); });
        for (int index = 1; index < chain.Length; index++) {
            Assert.Equal(chain[index - 1].Id, chain[index].PreviousFrameId);
            Assert.Equal(chain[index].Id, chain[index - 1].NextFrameId);
            Assert.Equal(chain[index - 1].TextStart + chain[index - 1].TextLength, chain[index].TextStart);
        }
        Assert.Equal(linked.Text, string.Concat(chain.Select(frame => linked.Text.Substring(frame.TextStart!.Value, frame.TextLength!.Value))));
        Assert.False(chain[2].HasOverflow);

        PublisherTextFrame expectedColumns = document.Pages[0].TextFrames.Single(frame => frame.Id == columns.Id);
        PublisherTextFrame actualColumns = Assert.Single(decoded.Pages[0].TextFrames, frame => frame.ColumnCount == 2);
        Assert.Null(actualColumns.PreviousFrameId); Assert.Null(actualColumns.NextFrameId);
        Assert.NotEqual(linked.Id, actualColumns.StoryId);
        Assert.Equal(Bounds(expectedColumns), Bounds(actualColumns));
        Assert.Equal((expectedColumns.ColumnCount, expectedColumns.ColumnSpacing, expectedColumns.Padding, expectedColumns.VerticalAlignment),
            (actualColumns.ColumnCount, actualColumns.ColumnSpacing, actualColumns.Padding, actualColumns.VerticalAlignment));
        AssertStory(document.TextStories.Single(story => story.Id == columns.StoryId),
            decoded.TextStories.Single(story => story.Id == actualColumns.StoryId));

        PublisherTable expectedTable = Assert.Single(document.Pages[2].Tables);
        PublisherTable actualTable = Assert.Single(decoded.Pages.SelectMany(page => page.Tables));
        Assert.Equal(decoded.Pages[2].Id, actualTable.PageId);
        Assert.True(actualTable.HasTextMapping);
        Assert.Equal((expectedTable.X, expectedTable.Y, expectedTable.Width, expectedTable.Height),
            (actualTable.X, actualTable.Y, actualTable.Width, actualTable.Height));
        Assert.Equal(expectedTable.ColumnWidths, actualTable.ColumnWidths);
        Assert.Equal(expectedTable.RowHeights, actualTable.RowHeights);
        Assert.Equal(expectedTable.Cells.Select(cell => (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan, cell.Text)),
            actualTable.Cells.Select(cell => (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan, cell.Text)));
        Assert.Empty(actualTable.BorderSegments);
    }

    [Fact]
    public void Creation_preserves_authored_none_wrapping_after_reload_and_allows_guarded_master_picture_edits() {
        PublisherDocument authored = Author(true);
        PublisherPicture authoredPicture = Assert.Single(authored.MasterPages.Single(page => page.MasterSide == PublisherMasterPageSide.Left).Pictures);
        authored.SetPictureBounds(authoredPicture.Id, 144, authoredPicture.Y, authoredPicture.Width, authoredPicture.Height);
        byte[] generated = PublisherNativeCreationCodec.Create(PublisherNativeCreationProjection.ComposedContent(authored),
            new PublisherReadOptions(), MaximumBytes, CancellationToken.None);
        PublisherDocument decoded = PublisherDocument.Load(generated);

        Assert.True(OfficeCompoundFileReader.TryRead(generated, out OfficeCompoundFile? compound, out string? error), error);
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        PublisherContents contents = new PublisherContentsReader(new PublisherBinaryData(compound!.Streams["Contents"], "Contents"), context).Read();
        PublisherEscherData artwork = new PublisherEscherReader(
            new PublisherBinaryData(compound.Streams["Escher/EscherStm"], "Escher/EscherStm"),
            compound.Streams.TryGetValue("Escher/EscherDelayStm", out byte[]? delay) ? new PublisherBinaryData(delay, "Escher/EscherDelayStm") : null,
            context).Read();
        uint[] semanticIds = decoded.MasterPages.Concat(decoded.Pages).SelectMany(page => page.ObjectIds).ToArray();
        var pictureIds = decoded.MasterPages.Concat(decoded.Pages).SelectMany(page => page.Pictures).Select(picture => picture.Id).ToHashSet();
        Assert.NotEmpty(semanticIds);
        foreach (uint id in semanticIds) {
            PublisherSourceShape source = contents.Shapes[id];
            PublisherEscherShape shape = Assert.Single(artwork.Shapes, item => item.Id == id);
            Assert.Equal(PublisherNativeWrapMode.None, PublisherWrapRelationships.ReadMode(source, shape,
                pictureIds.Contains(id), negativeMargins: false, tight: false));
            PublisherBlock mode = Assert.Single(source.Fields, field => field.Id == 4);
            Assert.Equal((0x10, 0x100U), ((int)mode.Type, mode.Value));
        }
        Assert.Equal(generated, decoded.ToPubBytes());

        PublisherPage left = decoded.MasterPages.Single(page => page.MasterSide == PublisherMasterPageSide.Left);
        PublisherPicture original = Assert.Single(left.Pictures);
        PublisherPicture[] pictures = decoded.MasterPages.Concat(decoded.Pages).SelectMany(page => page.Pictures).ToArray();
        var stories = decoded.TextStories.Select(story => (story.Id, story.Text)).ToArray();
        var owners = decoded.MasterPages.Concat(decoded.Pages).Select(page => (page.Id, page.MasterPageId, page.MasterSide, page.Name)).ToArray();
        decoded.SetPictureBounds(original.Id, original.X + 12, original.Y, original.Width, original.Height);

        OfficeConversionException refusal = Assert.Throws<OfficeConversionException>(() => decoded.ToPubBytes());
        PublisherSaveReport strict = Assert.IsType<PublisherSaveReport>(refusal.Report);
        Assert.False(strict.HasErrors);
        OfficeConversionFidelityDiagnostic qualification = Assert.Single(strict.FidelityDiagnostics, item => item.LossKind != OfficeConversionLossKind.None);
        Assert.Equal(("PUB_NATIVE_FIXED_PATCH_UNQUALIFIED", OfficeConversionLossKind.Unassessed), (qualification.Code, qualification.LossKind));
        var saved = decoded.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.False(saved.Report.HasErrors);
        PublisherDocument reopened = PublisherDocument.Load(saved.Value);
        Assert.Equal(owners, reopened.MasterPages.Concat(reopened.Pages).Select(page => (page.Id, page.MasterPageId, page.MasterSide, page.Name)));
        Assert.Equal(semanticIds, reopened.MasterPages.Concat(reopened.Pages).SelectMany(page => page.ObjectIds));
        Assert.Equal(stories, reopened.TextStories.Select(story => (story.Id, story.Text)));
        PublisherImage asset = Assert.Single(reopened.Images), expectedAsset = Assert.Single(decoded.Images);
        Assert.Equal(expectedAsset.Id, asset.Id);
        Assert.Equal(expectedAsset.GetBytes(), asset.GetBytes());
        PublisherPicture[] restored = reopened.MasterPages.Concat(reopened.Pages).SelectMany(page => page.Pictures).ToArray();
        Assert.Equal(pictures.Select(picture => (picture.Id, picture.PageId, picture.ImageId)),
            restored.Select(picture => (picture.Id, picture.PageId, picture.ImageId)));
        foreach (PublisherPicture picture in pictures) {
            PublisherPicture actual = restored.Single(item => item.Id == picture.Id);
            Assert.Equal((picture.X + (picture.Id == original.Id ? 12 : 0), picture.Y, picture.Width, picture.Height),
                (actual.X, actual.Y, actual.Width, actual.Height));
        }
        Assert.Equal(saved.Value, reopened.ToPubBytes());
    }

    [Fact]
    public void Composed_creation_refuses_three_centered_masters() {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        document.AddMasterPage("First");
        document.AddMasterPage("Second");
        document.AddMasterPage("Third");
        document.SetPageMaster(document.Pages[0].Id, document.MasterPages[0].Id);
        Assert.Throws<NotSupportedException>(() => PublisherNativeCreationProjection.ComposedContent(document));
    }

    private static PublisherDocument Author(bool paired) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        document.AddMasterPage("Centered brand");
        if (paired) document.AddMasterPagePair("Paired brand");
        else document.AddMasterPage("Alternate brand");
        int[] assignments = { 0, 1, paired ? 2 : 0 };
        document.SetPageMaster(document.Pages[0].Id, document.MasterPages[assignments[0]].Id);
        document.AddPage(document.MasterPages[assignments[1]].Id);
        document.AddPage(document.MasterPages[assignments[2]].Id);

        byte[] image = PublisherPictureEditingFixture.Image(OfficeColor.FromRgb(40, 120, 200));
        int? imageId = null;
        for (int index = 0; index < document.MasterPages.Count; index++) {
            uint masterId = document.MasterPages[index].Id;
            document.AddTextFrame(masterId, 36 + index * 12, 72, 360, 36, new[] { Paragraph("Master marker " + index) });
            OfficeShape rectangle = OfficeShape.Rectangle(240, 24);
            rectangle.FillColor = OfficeColor.FromRgb((byte)(60 + index * 60), 90, 150);
            rectangle.StrokeColor = null;
            document.AddShape(masterId, rectangle, 36 + index * 12, 36 + index * 6);
            PublisherPicture picture = imageId.HasValue
                ? document.AddPicture(masterId, imageId.Value, 504, 36, 36, 36)
                : document.AddPicture(masterId, image, 504, 36, 36, 36);
            imageId = picture.ImageId;
        }

        OfficeRichTextParagraph[] linkedParagraphs = {
            Paragraph(string.Concat(Enumerable.Repeat("The linked lead stays in its complete story. ", 12)).TrimEnd()),
            Paragraph(string.Concat(Enumerable.Repeat("The closing paragraph follows the reordered frames. ", 10)).TrimEnd())
        };
        PublisherTextFrame first = document.AddTextFrame(document.Pages[0].Id, 72, 144, 216, 36, linkedParagraphs);
        PublisherTextFrame middle = document.AddLinkedTextFrame(first.Id, document.Pages[2].Id, 72, 144, 216, 36);
        document.AddLinkedTextFrame(middle.Id, document.Pages[1].Id, 72, 144, 468, 360);
        PublisherTextFrame columns = document.AddTextFrame(document.Pages[0].Id, 72, 240, 360, 180,
            new[] { Paragraph("Independent columns retain their own story and frame layout.") });
        document.SetTextFrameLayout(columns.Id, 2, 18, new OfficeTextPadding(6, 9, 12, 3), OfficeTextVerticalAlignment.Top);
        PublisherTable table = document.AddTable(document.Pages[2].Id, 72, 600, new[] { 120D, 180D }, new[] { 42D });
        document.ReplaceTableCellParagraphs(table.Id, 0, 0, new[] { Paragraph("Item") });
        document.ReplaceTableCellParagraphs(table.Id, 0, 1, new[] { Paragraph("Count") });
        return document;
    }

    private static int[] MasterIndices(PublisherDocument document) {
        var indices = document.MasterPages.Select((page, index) => (page.Id, index)).ToDictionary(item => item.Id, item => item.index);
        return document.Pages.Select(page => indices[page.MasterPageId!.Value]).ToArray();
    }

    private static (double X, double Y, double Width, double Height) Bounds(PublisherTextFrame frame) =>
        (frame.X, frame.Y, frame.Width, frame.Height);

    private static OfficeRichTextParagraph Paragraph(string text) => new(new[] {
        new OfficeRichTextRun(text, 12, OfficeColor.Black, fontFamily: "Arial").WithLanguageId(1033)
    }, lineHeight: 18);

    private static void AssertStory(PublisherTextStory expected, PublisherTextStory actual) {
        Assert.Equal(expected.Text + "\n", actual.Text);
        Assert.Equal(expected.Paragraphs.Count, actual.Paragraphs.Count);
        for (int index = 0; index < expected.Paragraphs.Count; index++) {
            Assert.Equal(expected.Paragraphs[index].LineHeight, actual.Paragraphs[index].LineHeight);
            Assert.Equal(expected.Paragraphs[index].Runs.Select(run => (run.Text, run.FontSize, run.Color, run.FontFamily, run.LanguageId)),
                actual.Paragraphs[index].Runs.Select(run => (run.Text, run.FontSize, run.Color, run.FontFamily, run.LanguageId)));
        }
    }
}
