using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherMasterContentTests {
    [Fact]
    public void Master_content_is_owned_once_and_edits_reproject_assigned_pages_without_changing_old_scenes() {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Brand");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        document.AddPage(master.Id);
        PublisherPage last = document.AddPage(master.Id);
        var paint = OfficeShape.Rectangle(240, 96);
        paint.FillColor = OfficeColor.Red; paint.StrokeColor = null;
        PublisherShape background = document.AddShape(master.Id, paint, 36, 72);
        byte[] original = PublisherPictureEditingFixture.Image(OfficeColor.Blue);
        byte[] replacement = PublisherPictureEditingFixture.Image(OfficeColor.White);
        PublisherPicture logo = document.AddPicture(master.Id, original, 72, 72, 96, 96);
        PublisherTextFrame title = document.AddTextFrame(master.Id, 36, 72, 288, 108, Paragraph("Master content"));
        PublisherPicture local = document.AddPicture(last.Id, logo.ImageId!.Value, 180, 72, 96, 96);
        var oldPages = document.Pages; var oldMaster = document.MasterPages.Single();
        document.ReplaceStoryParagraphs(title.StoryId, Paragraph("Updated master"));
        document.SetShapeBounds(background.Id, 36, 60, 240, 96);
        document.ReplacePictureImage(logo.Id, replacement);
        Assert.Equal(3, Owned(document.MasterPages.Single()).Length);
        Assert.Equal(new[] { 0, 0, 1 }, document.Pages.Select(page => Owned(page).Length));
        Assert.Equal(4, document.ReadReport.CurrentObjectCount);
        Assert.Single(document.TextStories);
        Assert.Equal(2, document.Images.Count);
        Assert.Equal(local.ImageId, document.Pages[2].Pictures.Single().ImageId);
        Assert.NotEqual(document.MasterPages.Single().Pictures.Single().ImageId, local.ImageId);
        for (int index = 0; index < 3; index++) {
            OfficeDrawing drawing = document.Pages[index].CreateDrawing();
            var images = PublisherNativeTests.Elements(drawing).OfType<OfficeDrawingImage>().ToArray();
            Assert.Equal(index == 2 ? 2 : 1, images.Length);
            Assert.Equal(replacement, images[0].Bytes);
            if (index == 2) Assert.Equal(original, images[1].Bytes);
            Assert.Contains(PublisherNativeTests.Elements(drawing).OfType<OfficeDrawingRichText>(), text => PublisherNativeTests.Text(text) == "Updated master");
            Assert.Contains(PublisherNativeTests.Elements(oldPages[index].CreateDrawing()).OfType<OfficeDrawingRichText>(), text => PublisherNativeTests.Text(text) == "Master content");
        }
        Assert.Equal(72, oldMaster.Shapes.Single().Y);
        Assert.Equal(original, PublisherNativeTests.Elements(oldMaster.CreateDrawing()).OfType<OfficeDrawingImage>().Single().Bytes);
    }

    [Theory]
    [InlineData("MasterContent", false)]
    [InlineData("MasterContentEdited", true)]
    public void Native_populated_master_keeps_definition_and_page_ownership_separate(string fixture, bool edited) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture($"NativeComparison/{fixture}/native-{fixture.ToLowerInvariant()}.pub"));
        PublisherPage master = Assert.Single(document.MasterPages);
        Assert.All(document.Pages, page => Assert.Equal(master.Id, page.MasterPageId));
        Assert.Equal(3, Owned(master).Length);
        Assert.Equal(new[] { 1, 0, 3 }, document.Pages.Select(page => Owned(page).Length));
        Assert.Equal(7, document.ReadReport.CurrentObjectCount);
        Assert.Equal(edited ? 2 : 1, document.Images.Count);
        Assert.Equal(edited ? 2 : 1, Assert.Single(master.Pictures).ImageId);
        Assert.Equal(1, Assert.Single(document.Pages[2].Pictures).ImageId);
        Assert.Equal(2, document.MasterPages.Concat(document.Pages).Sum(page => page.Pictures.Count));
        for (int index = 0; index < document.Pages.Count; index++) {
            var elements = PublisherNativeTests.Elements(document.Pages[index].CreateDrawing()).ToArray();
            Assert.Equal(index == 2 ? 2 : 1, elements.OfType<OfficeDrawingImage>().Count());
            string expected = edited ? "Updated master content" : "Master: shared content";
            Assert.Contains(elements.OfType<OfficeDrawingRichText>(), text => PublisherNativeTests.Text(text) == expected);
        }
    }

    private static uint[] Owned(PublisherPage page) => page.Shapes.Select(shape => shape.Id).Concat(page.Pictures.Select(picture => picture.Id))
        .Concat(page.TextFrames.Select(frame => frame.Id)).Distinct().ToArray();
    private static OfficeRichTextParagraph[] Paragraph(string text) => new[] {
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, 18, OfficeColor.Black, fontFamily: "Arial") })
    };
}
