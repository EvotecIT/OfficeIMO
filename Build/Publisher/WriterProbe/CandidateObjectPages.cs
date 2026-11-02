using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects a closed source-free shape/picture specimen without manufacturing a semantic text story.</summary>
internal static class CandidateObjectPages {
    internal static CandidatePublicationSpec Create(byte[] image) =>
        CandidateModelProjection.SingleMasterContent(Author(image), includePictureNames: false);

    internal static PublisherDocument Author(byte[] image) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        var rectangle = OfficeShape.Rectangle(144, 72);
        rectangle.FillColor = OfficeColor.FromRgb(255, 0, 0);
        rectangle.StrokeColor = null;
        document.AddShape(document.Pages[0].Id, rectangle, 72, 216);
        document.AddPage(master.Id);
        PublisherPage last = document.AddPage(master.Id);
        document.AddPicture(last.Id, image, 330, 216, 96, 96);
        return document;
    }
}
