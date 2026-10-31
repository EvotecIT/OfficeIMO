using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects a closed source-free shape/picture specimen without manufacturing a semantic text story.</summary>
internal static class CandidateObjectPages {
    internal static CandidatePublicationSpec Create(byte[] image) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        var rectangle = OfficeShape.Rectangle(144, 72);
        rectangle.FillColor = OfficeColor.FromRgb(255, 0, 0);
        rectangle.StrokeColor = null;
        PublisherShape shape = document.AddShape(document.Pages[0].Id, rectangle, 72, 216);
        document.AddPage(master.Id);
        PublisherPage last = document.AddPage(master.Id);
        PublisherPicture picture = document.AddPicture(last.Id, image, 330, 216, 96, 96);

        var objects = new[] {
            new CandidateObjectSpec(CandidateObjectKind.Rectangle,
                new CandidatePlacement(shape.X, shape.Y, shape.Width, shape.Height), Paint: shape.CreateShape()),
            new CandidateObjectSpec(CandidateObjectKind.Picture,
                new CandidatePlacement(picture.X, picture.Y, picture.Width, picture.Height),
                Image: document.Images.Single(asset => asset.Id == picture.ImageId).GetBytes(),
                PageIndex: document.Pages.Count - 1, Paint: picture.CreatePaint())
        };
        return new CandidatePublicationSpec(checked((uint)(document.Pages[0].Width * 12700)),
            checked((uint)(document.Pages[0].Height * 12700)), objects, document.Pages.Count);
    }
}
