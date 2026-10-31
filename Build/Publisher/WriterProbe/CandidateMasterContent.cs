using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Authors one populated master and page-owned foreground without duplicating the master objects.</summary>
internal static class CandidateMasterContent {
    internal static CandidatePublicationSpec Create(byte[] image) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        document.AddPage(master.Id);
        document.AddPage(master.Id);
        var objects = new List<CandidateObjectSpec>();
        Shape(master.Id, 0, true, 36, 72, 240, 96, OfficeColor.Red);
        PublisherPicture logo = Picture(master.Id, 0, true, 72, null);
        Text(master.Id, 0, true, "Master: shared content", 36, 72, 18);
        Shape(document.Pages[0].Id, 0, false, 36, 84, 240, 12, OfficeColor.Blue);
        Picture(document.Pages[2].Id, 2, false, 180, logo.ImageId);
        Shape(document.Pages[2].Id, 2, false, 36, 108, 240, 12, OfficeColor.FromRgb(0, 128, 0));
        Text(document.Pages[2].Id, 2, false, "Page foreground", 36, 180, 12);
        return new CandidatePublicationSpec(612 * 12700, 792 * 12700, objects, document.Pages.Count);

        void Shape(uint owner, int index, bool onMaster, double x, double y, double width, double height, OfficeColor color) {
            var paint = OfficeShape.Rectangle(width, height);
            paint.FillColor = color; paint.StrokeColor = null;
            PublisherShape shape = document.AddShape(owner, paint, x, y);
            objects.Add(new CandidateObjectSpec(CandidateObjectKind.Rectangle,
                new CandidatePlacement(shape.X, shape.Y, shape.Width, shape.Height),
                PageIndex: index, Paint: shape.CreateShape(), OnMaster: onMaster));
        }
        PublisherPicture Picture(uint owner, int index, bool onMaster, double x, int? imageId) {
            PublisherPicture picture = imageId.HasValue
                ? document.AddPicture(owner, imageId.Value, x, 72, 96, 96)
                : document.AddPicture(owner, image, x, 72, 96, 96);
            objects.Add(new CandidateObjectSpec(CandidateObjectKind.Picture,
                new CandidatePlacement(picture.X, picture.Y, picture.Width, picture.Height),
                Image: document.Images.Single(asset => asset.Id == picture.ImageId).GetBytes(),
                PageIndex: index, Paint: picture.CreatePaint(), ImageName: "Image1.png", ImageId: picture.ImageId, OnMaster: onMaster));
            return picture;
        }
        void Text(uint owner, int index, bool onMaster, string text, double x, double y, double size) {
            PublisherTextFrame frame = document.AddTextFrame(owner, x, y, 288, 108, new[] {
                new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black, fontFamily: "Arial") })
            });
            objects.Add(new CandidateObjectSpec(CandidateObjectKind.TextBox,
                new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height),
                CandidateStorySpec.Styled(document.TextStories.Single(story => story.Id == frame.StoryId).Paragraphs.Single()),
                PageIndex: index, ColumnSpacing: frame.ColumnSpacing, OnMaster: onMaster));
        }
    }
}
