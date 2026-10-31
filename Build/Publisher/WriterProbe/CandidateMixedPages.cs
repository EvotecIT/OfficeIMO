using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Authors overlapping object kinds in different append orders on separated physical pages.</summary>
internal static class CandidateMixedPages {
    internal static CandidatePublicationSpec Create(byte[] image) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        document.AddPage(master.Id);
        document.AddPage(master.Id);
        var objects = new List<CandidateObjectSpec>();
        Shape(0, 72, 144, 240, 120, OfficeColor.Red);
        PublisherPicture first = Picture(0, 108, null);
        Text(0, "Mixed first: overlap order");
        Shape(0, 72, 168, 240, 12, OfficeColor.Blue);
        Text(2, "Mixed last: overlap order");
        Picture(2, 108, first.ImageId);
        Shape(2, 72, 168, 240, 12, OfficeColor.FromRgb(0, 128, 0));
        Picture(2, 216, first.ImageId);
        return new CandidatePublicationSpec(612 * 12700, 792 * 12700, objects, document.Pages.Count);

        void Shape(int pageIndex, double x, double y, double width, double height, OfficeColor color) {
            var paint = OfficeShape.Rectangle(width, height);
            paint.FillColor = color; paint.StrokeColor = null;
            PublisherShape shape = document.AddShape(document.Pages[pageIndex].Id, paint, x, y);
            objects.Add(new CandidateObjectSpec(CandidateObjectKind.Rectangle,
                new CandidatePlacement(shape.X, shape.Y, shape.Width, shape.Height),
                PageIndex: pageIndex, Paint: shape.CreateShape()));
        }
        PublisherPicture Picture(int pageIndex, double x, int? imageId) {
            uint pageId = document.Pages[pageIndex].Id;
            PublisherPicture picture = imageId.HasValue
                ? document.AddPicture(pageId, imageId.Value, x, 144, 96, 96)
                : document.AddPicture(pageId, image, x, 144, 96, 96);
            objects.Add(new CandidateObjectSpec(CandidateObjectKind.Picture,
                new CandidatePlacement(picture.X, picture.Y, picture.Width, picture.Height),
                Image: document.Images.Single(asset => asset.Id == picture.ImageId).GetBytes(),
                PageIndex: pageIndex, Paint: picture.CreatePaint(), ImageName: "Image1.png", ImageId: picture.ImageId));
            return picture;
        }
        void Text(int pageIndex, string text) {
            var paragraph = new OfficeRichTextParagraph(new[] {
                new OfficeRichTextRun(text, 18, OfficeColor.Black, fontFamily: "Arial")
            });
            PublisherTextFrame frame = document.AddTextFrame(document.Pages[pageIndex].Id, 72, 144, 288, 108,
                new[] { paragraph });
            objects.Add(new CandidateObjectSpec(CandidateObjectKind.TextBox,
                new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height),
                CandidateStorySpec.Styled(document.TextStories.Single(story => story.Id == frame.StoryId).Paragraphs.Single()),
                PageIndex: pageIndex, ColumnSpacing: frame.ColumnSpacing));
        }
    }
}
