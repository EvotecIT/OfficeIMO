using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects a closed source-free shape/picture specimen without manufacturing a semantic text story.</summary>
internal static class CandidateObjectPages {
    internal static PublisherNativeCreationPlan Create(byte[] image) =>
        PublisherNativeCreationProjection.SingleMasterContent(Author(image), includePictureNames: false);

    internal static PublisherDocument Author(byte[] image) {
        PublisherDocument document = PublisherDocument.Create(OfficeLength.Points(612), OfficeLength.Points(792));
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        var rectangle = OfficeShape.Rectangle(144, 72);
        rectangle.FillColor = OfficeColor.FromRgb(255, 0, 0);
        rectangle.StrokeColor = null;
        document.AddShape(document.Pages[0].Id, rectangle, OfficeLength.Points(72), OfficeLength.Points(216));
        document.AddPage(master.Id);
        PublisherPage last = document.AddPage(master.Id);
        document.AddImage(last.Id, image, OfficeLength.Points(330), OfficeLength.Points(216), OfficeLength.Points(96), OfficeLength.Points(96));
        return document;
    }
}
