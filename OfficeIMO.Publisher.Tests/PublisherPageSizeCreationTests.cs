using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPageSizeCreationTests {
    private const long MaximumBytes = 64L * 1024 * 1024;

    [Theory]
    [InlineData(7560000U, 10692000U)]
    [InlineData(3657600U, 5486400U)]
    [InlineData(5486400U, 3657600U)]
    public void Exact_page_sizes_keep_mixed_content_ownership_guides_and_none_wrapping(uint widthEmu, uint heightEmu) {
        double width = widthEmu / 12700D, height = heightEmu / 12700D;
        PublisherDocument authored = Author(width, height);
        byte[] generated = PublisherNativeCreationCodec.Create(PublisherNativeCreationProjection.ComposedContent(authored),
            new PublisherReadOptions(), MaximumBytes, CancellationToken.None);
        PublisherDocument decoded = PublisherDocument.Load(generated);

        Assert.Equal(2, decoded.Pages.Count);
        Assert.Equal(authored.TextStories.Count, decoded.TextStories.Count);
        PublisherPage master = Assert.Single(decoded.MasterPages);
        Assert.Equal((authored.MasterPages[0].Name, PublisherMasterPageSide.Single), (master.Name, master.MasterSide!.Value));
        Assert.Null(master.MasterPageId);
        Assert.All(decoded.Pages, page => { Assert.Equal(master.Id, page.MasterPageId); Assert.Null(page.MasterSide); });
        PublisherImage asset = Assert.Single(decoded.Images);
        Assert.Equal(Assert.Single(authored.Images).GetBytes(), asset.GetBytes());
        Assert.Equal("image/png", asset.ContentType);
        Assert.Equal(2, decoded.MasterPages.Concat(decoded.Pages).Sum(page => page.Pictures.Count));
        var expectedOwners = authored.MasterPages.Concat(authored.Pages).ToArray();
        var actualOwners = decoded.MasterPages.Concat(decoded.Pages).ToArray();
        for (int index = 0; index < expectedOwners.Length; index++) {
            PublisherPage expected = expectedOwners[index], actual = actualOwners[index];
            Assert.Equal(width, actual.Width, 6); Assert.Equal(height, actual.Height, 6);
            Assert.Equal(expected.ObjectIds.Count, actual.ObjectIds.Count);
            Assert.Equal(expected.TextFrames.Count, actual.TextFrames.Count);
            for (int frameIndex = 0; frameIndex < expected.TextFrames.Count; frameIndex++) {
                PublisherTextFrame before = expected.TextFrames[frameIndex], after = actual.TextFrames[frameIndex];
                Assert.Equal(actual.Id, after.PageId);
                AssertBounds((before.X, before.Y, before.Width, before.Height), (after.X, after.Y, after.Width, after.Height));
                Assert.Equal((before.ColumnCount, before.ColumnSpacing, before.Padding, before.VerticalAlignment),
                    (after.ColumnCount, after.ColumnSpacing, after.Padding, after.VerticalAlignment));
                PublisherTextStory story = authored.TextStories.Single(item => item.Id == before.StoryId);
                PublisherTextStory restored = decoded.TextStories.Single(item => item.Id == after.StoryId);
                Assert.Equal(story.Text + "\n", restored.Text);
                Assert.Equal(story.Paragraphs.Count, restored.Paragraphs.Count);
                Assert.Equal(story.Paragraphs.Select(paragraph => paragraph.LineHeight), restored.Paragraphs.Select(paragraph => paragraph.LineHeight));
                for (int paragraphIndex = 0; paragraphIndex < story.Paragraphs.Count; paragraphIndex++)
                    Assert.Equal(story.Paragraphs[paragraphIndex].Runs.Select(run => (run.Text, run.FontSize, run.Bold, run.Color, run.FontFamily, run.LanguageId)),
                        restored.Paragraphs[paragraphIndex].Runs.Select(run => (run.Text, run.FontSize, run.Bold, run.Color, run.FontFamily, run.LanguageId)));
            }
            PublisherShape[] rectangles = actual.Shapes.Where(shape => actual.TextFrames.All(frame => frame.Id != shape.Id)).ToArray();
            Assert.Equal(expected.Shapes.Count, rectangles.Length);
            for (int shapeIndex = 0; shapeIndex < rectangles.Length; shapeIndex++) {
                PublisherShape before = expected.Shapes[shapeIndex], after = rectangles[shapeIndex];
                Assert.Equal(actual.Id, after.PageId);
                AssertBounds((before.X, before.Y, before.Width, before.Height), (after.X, after.Y, after.Width, after.Height));
                Assert.Equal((OfficeShapeKind.Rectangle, before.CreateShape().FillColor), (after.CreateShape().Kind, after.CreateShape().FillColor));
            }
            Assert.Equal(expected.Pictures.Count, actual.Pictures.Count);
            for (int pictureIndex = 0; pictureIndex < expected.Pictures.Count; pictureIndex++) {
                PublisherPicture before = expected.Pictures[pictureIndex], after = actual.Pictures[pictureIndex];
                Assert.Equal(actual.Id, after.PageId); Assert.Equal(asset.Id, after.ImageId);
                AssertBounds((before.X, before.Y, before.Width, before.Height), (after.X, after.Y, after.Width, after.Height));
            }
        }
        uint[] owned = actualOwners.SelectMany(page => page.ObjectIds).ToArray();
        Assert.Equal(owned.Length, owned.Distinct().Count());
        PublisherTextFrame[] linked = decoded.Pages.SelectMany(page => page.TextFrames)
            .Where(frame => frame.PreviousFrameId.HasValue || frame.NextFrameId.HasValue).OrderBy(frame => frame.Order).ToArray();
        Assert.Equal(2, linked.Length);
        Assert.Equal(linked[0].StoryId, linked[1].StoryId);
        Assert.Equal(new uint[] { 0, 1 }, linked.Select(frame => frame.Order));
        Assert.Equal(decoded.Pages.Select(page => page.Id), linked.Select(frame => frame.PageId));
        Assert.Equal(linked[1].Id, linked[0].NextFrameId); Assert.Equal(linked[0].Id, linked[1].PreviousFrameId);
        Assert.Null(linked[0].PreviousFrameId); Assert.Null(linked[1].NextFrameId);
        Assert.Equal(linked[0].TextStart + linked[0].TextLength, linked[1].TextStart);
        PublisherTable expectedTable = Assert.Single(authored.Pages[1].Tables);
        PublisherTable table = Assert.Single(decoded.Pages.SelectMany(page => page.Tables));
        Assert.Empty(master.Tables); Assert.Equal(decoded.Pages[1].Id, table.PageId);
        Assert.True(table.HasTextMapping);
        AssertBounds((expectedTable.X, expectedTable.Y, expectedTable.Width, expectedTable.Height), (table.X, table.Y, table.Width, table.Height));
        Assert.Equal(expectedTable.ColumnWidths, table.ColumnWidths); Assert.Equal(expectedTable.RowHeights, table.RowHeights);
        Assert.Equal(expectedTable.Cells.Select(cell => (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan, cell.Text)),
            table.Cells.Select(cell => (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan, cell.Text)));
        Assert.Empty(table.BorderSegments);

        Assert.True(OfficeCompoundFileReader.TryRead(generated, out OfficeCompoundFile? compound, out string? error), error);
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        var data = new PublisherBinaryData(compound!.Streams["Contents"], "Contents");
        var blocks = new PublisherBlockReader(data, context);
        PublisherContents contents = new PublisherContentsReader(data, context).Read();
        PublisherContentChunk document = Assert.Single(contents.Chunks.Values, chunk => chunk.Kind == 0x44);
        var dimensions = blocks.Children(blocks.Chunk(document.Offset, document.End).Single(field => field.Id == 0x12));
        Assert.Equal((widthEmu, heightEmu), (dimensions.Single(field => field.Id == 1).Value, dimensions.Single(field => field.Id == 2).Value));
        foreach (PublisherPage definition in decoded.MasterPages) {
            PublisherContentChunk descriptor = contents.Chunks[definition.Id];
            uint geometryId = blocks.Chunk(descriptor.Offset, descriptor.End).Single(field => field.Id == 3).Value;
            PublisherContentChunk geometry = contents.Chunks[geometryId];
            var fields = blocks.Chunk(geometry.Offset, geometry.End);
            var extent = blocks.Children(fields.Single(field => field.Id == 3));
            Assert.Equal((widthEmu, heightEmu), (extent.Single(field => field.Id == 1).Value, extent.Single(field => field.Id == 2).Value));
            Assert.Equal(new uint[] { 0, widthEmu, 0, heightEmu }, blocks.Children(fields.Single(field => field.Id == 2))
                .Select(guide => blocks.Children(guide).Single(field => field.Id == 1).Value));
        }
        PublisherEscherData artwork = new PublisherEscherReader(
            new PublisherBinaryData(compound.Streams["Escher/EscherStm"], "Escher/EscherStm"),
            compound.Streams.TryGetValue("Escher/EscherDelayStm", out byte[]? delay) ? new PublisherBinaryData(delay, "Escher/EscherDelayStm") : null, context).Read();
        var pictures = actualOwners.SelectMany(page => page.Pictures).Select(picture => picture.Id).ToHashSet();
        foreach (uint id in owned) {
            PublisherSourceShape source = contents.Shapes[id];
            Assert.Equal(PublisherNativeWrapMode.None, PublisherWrapRelationships.ReadMode(source,
                Assert.Single(artwork.Shapes, shape => shape.Id == id), pictures.Contains(id), negativeMargins: false, tight: false));
            PublisherBlock mode = Assert.Single(source.Fields, field => field.Id == 4);
            Assert.Equal((0x10, 0x100U), ((int)mode.Type, mode.Value));
        }
    }

    [Fact]
    public void Creation_preserves_unassigned_physical_content_and_retained_master_suppression() {
        PublisherDocument authored = Author(288, 432);
        byte[] appliedBytes = PublisherNativeCreationCodec.Create(PublisherNativeCreationProjection.ComposedContent(authored),
            new PublisherReadOptions(), MaximumBytes, CancellationToken.None);
        PublisherDocument applied = PublisherDocument.Load(appliedBytes);
        authored.SetPageMaster(authored.Pages[1].Id, null);
        byte[] clearedBytes = PublisherNativeCreationCodec.Create(PublisherNativeCreationProjection.ComposedContent(authored),
            new PublisherReadOptions(), MaximumBytes, CancellationToken.None);
        PublisherDocument cleared = PublisherDocument.Load(clearedBytes);

        PublisherPage master = Assert.Single(cleared.MasterPages);
        Assert.False(authored.HasNativeSource);
        Assert.Equal(master.Id, cleared.Pages[0].MasterPageId);
        Assert.Null(cleared.Pages[1].MasterPageId);
        Assert.Equal(new[] { master.Id, cleared.Pages[0].Id, cleared.Pages[1].Id },
            applied.MasterPages.Concat(applied.Pages).Select(page => page.Id));
        foreach (var pair in applied.MasterPages.Concat(applied.Pages).Zip(cleared.MasterPages.Concat(cleared.Pages), (before, after) => (before, after))) {
            Assert.Equal((288D, 432D), (pair.after.Width, pair.after.Height));
            Assert.Equal(pair.before.ObjectIds, pair.after.ObjectIds);
        }
        Assert.Equal(applied.TextStories.Select(story => (story.Id, story.Text)), cleared.TextStories.Select(story => (story.Id, story.Text)));
        Assert.Equal(Assert.Single(applied.Images).GetBytes(), Assert.Single(cleared.Images).GetBytes());
        Assert.Equal(Assert.Single(applied.Pages[1].Tables).Cells.Select(cell => cell.Text), Assert.Single(cleared.Pages[1].Tables).Cells.Select(cell => cell.Text));
        Assert.Contains("Sized publication", PublisherAuthoringTests.PaintedText(cleared.Pages[0]));
        Assert.DoesNotContain("Sized publication", PublisherAuthoringTests.PaintedText(cleared.Pages[1]));
        Assert.Contains("Item", PublisherAuthoringTests.PaintedText(cleared.Pages[1]));
        uint logoId = Assert.Single(master.Pictures).Id;
        uint ruleId = Assert.Single(master.Shapes, shape => master.TextFrames.All(frame => frame.Id != shape.Id)).Id;
        Assert.Contains(PublisherNativeTests.Elements(cleared.Pages[0].CreateDrawing()).OfType<OfficeDrawingImage>(),
            image => image.SourceElementIds?.Contains("publisher-object-" + logoId) == true);
        Assert.DoesNotContain(PublisherNativeTests.Elements(cleared.Pages[1].CreateDrawing()).OfType<OfficeDrawingImage>(),
            image => image.SourceElementIds?.Contains("publisher-object-" + logoId) == true);
        Assert.Contains(PublisherNativeTests.Elements(cleared.Pages[0].CreateDrawing()).OfType<OfficeDrawingShape>(),
            shape => shape.SourceElementIds?.Contains("publisher-object-" + ruleId) == true);
        Assert.DoesNotContain(PublisherNativeTests.Elements(cleared.Pages[1].CreateDrawing()).OfType<OfficeDrawingShape>(),
            shape => shape.SourceElementIds?.Contains("publisher-object-" + ruleId) == true);
        uint physicalPictureId = Assert.Single(cleared.Pages[1].Pictures).Id;
        Assert.Contains(PublisherNativeTests.Elements(cleared.Pages[1].CreateDrawing()).OfType<OfficeDrawingImage>(),
            image => image.SourceElementIds?.Contains("publisher-object-" + physicalPictureId) == true);

        Assert.True(OfficeCompoundFileReader.TryRead(appliedBytes, out OfficeCompoundFile? appliedCompound, out string? appliedError), appliedError);
        Assert.True(OfficeCompoundFileReader.TryRead(clearedBytes, out OfficeCompoundFile? clearedCompound, out string? clearedError), clearedError);
        foreach (var stream in appliedCompound!.Streams.Where(stream => stream.Key != "Contents"))
            Assert.Equal(stream.Value, clearedCompound!.Streams[stream.Key]);
        byte[] beforeContents = appliedCompound.Streams["Contents"], afterContents = clearedCompound!.Streams["Contents"];
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        var data = new PublisherBinaryData(afterContents, "Contents");
        PublisherContents native = new PublisherContentsReader(data, context).Read();
        PublisherContents beforeNative = new PublisherContentsReader(new PublisherBinaryData(beforeContents, "Contents"),
            new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None)).Read();
        uint clearedPageId = cleared.Pages[1].Id;
        PublisherSourcePage nativePage = native.Pages.Single(page => page.Id == clearedPageId);
        Assert.Equal(master.Id, nativePage.Master);
        Assert.False(nativePage.DisplayMaster);
        PublisherContentChunk descriptor = native.Chunks[clearedPageId];
        var fields = new PublisherBlockReader(data, context).Chunk(descriptor.Offset, descriptor.End);
        PublisherBlock display = Assert.Single(fields, field => field.Id == 7);
        Assert.Equal((byte)8, display.Type); Assert.Equal(0, display.Length);
        Assert.Equal((byte)7, afterContents[display.Offset - 2]);
        Assert.Equal((byte)8, afterContents[display.Offset - 1]);
        PublisherBlock reference = Assert.Single(fields, field => field.Id == 0x0D);
        Assert.Equal((byte)0x68, reference.Type); Assert.Equal(master.Id, reference.Value);
        foreach (PublisherContentChunk before in beforeNative.Chunks.Values.Where(chunk => chunk.Id != clearedPageId)) {
            PublisherContentChunk after = native.Chunks[before.Id];
            Assert.Equal((before.Kind, before.Parent), (after.Kind, after.Parent));
            Assert.Equal(beforeContents.Skip(before.Offset).Take(before.End - before.Offset).ToArray(),
                afterContents.Skip(after.Offset).Take(after.End - after.Offset).ToArray());
        }

        cleared.SetPageMaster(clearedPageId, master.Id);
        PublisherDocument restored = PublisherDocument.Load(cleared.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.All(restored.Pages, page => Assert.Equal(master.Id, page.MasterPageId));
        Assert.Contains("Sized publication", PublisherAuthoringTests.PaintedText(restored.Pages[1]));
        Assert.Equal(applied.TextStories.Select(story => (story.Id, story.Text)), restored.TextStories.Select(story => (story.Id, story.Text)));
    }

    [Theory]
    [InlineData("signed-range")]
    [InlineData("fractional-EMU")]
    public void Native_projection_refuses_unrepresentable_page_dimensions_without_mutating_the_model(string boundary) {
        double width = boundary == "signed-range" ? ((double)int.MaxValue + 1) / 12700 : 612 + 0.5 / 12700;
        PublisherDocument document = EmptyPublication(width, 792);
        AssertRefusedUnchanged(document, () => PublisherNativeCreationProjection.ComposedContent(document));
    }

    [Theory]
    [InlineData("odd-centered")]
    [InlineData("fractional-final")]
    [InlineData("overflow-final")]
    public void Native_creation_refuses_unrepresentable_complete_anchor_edges_without_mutating_the_model(string boundary) {
        double width = boundary == "odd-centered" ? (612D * 12700 + 1) / 12700 : 612;
        PublisherDocument document = EmptyPublication(width, 792);
        double x = boundary == "overflow-final" ? 306 + (int.MaxValue - 12700) / 12700D : 36;
        double objectWidth = boundary == "fractional-final" ? 12 + 0.5 / 12700 : boundary == "overflow-final" ? 2 : 12;
        var rectangle = OfficeShape.Rectangle(objectWidth, 12);
        rectangle.FillColor = OfficeColor.Blue; rectangle.StrokeColor = null;
        document.AddShape(document.Pages[0].Id, rectangle, x, 36);
        AssertRefusedUnchanged(document, () => PublisherNativeCreationCodec.Create(PublisherNativeCreationProjection.ComposedContent(document),
            new PublisherReadOptions(), MaximumBytes, CancellationToken.None));
    }

    private static PublisherDocument Author(double width, double height) {
        PublisherDocument document = EmptyPublication(width, height);
        uint master = document.MasterPages[0].Id;
        document.AddPage(master);
        document.AddTextFrame(master, 18, 18, width - 84, 24, new[] { Paragraph("Sized publication", true) });
        var rule = OfficeShape.Rectangle(width - 36, 3);
        rule.FillColor = OfficeColor.Blue; rule.StrokeColor = null;
        document.AddShape(master, rule, 18, height - 18);
        PublisherPicture logo = document.AddPicture(master, PublisherPictureEditingFixture.Image(OfficeColor.FromRgb(40, 120, 200)), width - 42, 18, 24, 24);
        PublisherTextFrame first = document.AddTextFrame(document.Pages[0].Id, 18, 54, width - 36, 36, new[] {
            Paragraph("A dimensioned publication keeps its complete linked story.", true),
            Paragraph(string.Concat(Enumerable.Repeat("Continuation content remains separate from page geometry. ", 5)).TrimEnd())
        });
        document.AddLinkedTextFrame(first.Id, document.Pages[1].Id, 18, 54, width - 36, 72);
        PublisherTextFrame columns = document.AddTextFrame(document.Pages[0].Id, 18, height - 162, width - 36, 66,
            new[] { Paragraph("Independent columns retain their own layout.") });
        document.SetTextFrameLayout(columns.Id, 2, 18, new OfficeTextPadding(6, 9, 12, 3), OfficeTextVerticalAlignment.Top);
        document.AddPicture(document.Pages[1].Id, logo.ImageId!.Value, width - 54, height - 42, 36, 36);
        PublisherTable table = document.AddTable(document.Pages[1].Id, width - 174, height - 90, new[] { 72D, 84D }, new[] { 30D });
        document.ReplaceTableCellParagraphs(table.Id, 0, 0, new[] { Paragraph("Item") });
        document.ReplaceTableCellParagraphs(table.Id, 0, 1, new[] { Paragraph("Count") });
        return document;
    }

    private static PublisherDocument EmptyPublication(double width, double height) {
        PublisherDocument document = PublisherDocument.Create(width, height);
        PublisherPage master = document.AddMasterPage("Page-size brand");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        return document;
    }

    private static OfficeRichTextParagraph Paragraph(string text, bool bold = false) => new(new[] {
        new OfficeRichTextRun(text, 12, OfficeColor.Black, bold: bold, fontFamily: "Arial").WithLanguageId(1033)
    }, lineHeight: 18);

    private static void AssertBounds((double X, double Y, double Width, double Height) expected, (double X, double Y, double Width, double Height) actual) {
        Assert.Equal(expected.X, actual.X, 6); Assert.Equal(expected.Y, actual.Y, 6);
        Assert.Equal(expected.Width, actual.Width, 6); Assert.Equal(expected.Height, actual.Height, 6);
    }

    private static void AssertRefusedUnchanged(PublisherDocument document, Action create) {
        var pages = document.Pages; var masters = document.MasterPages;
        var stories = document.TextStories; var images = document.Images; var report = document.ReadReport;
        Assert.Throws<NotSupportedException>(create);
        Assert.Same(pages, document.Pages); Assert.Same(masters, document.MasterPages);
        Assert.Same(stories, document.TextStories); Assert.Same(images, document.Images); Assert.Same(report, document.ReadReport);
    }
}
