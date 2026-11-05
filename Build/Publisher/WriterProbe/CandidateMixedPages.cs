using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Authors overlapping object kinds in different append orders on separated physical pages.</summary>
internal static class CandidateMixedPages {
    internal static PublisherNativeCreationPlan Create(byte[] image) => PublisherNativeCreationProjection.SingleMasterContent(Author(image));

    internal static PublisherDocument Author(byte[] image) {
        PublisherDocument document = PublisherDocument.Create(OfficeLength.Points(612), OfficeLength.Points(792));
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        document.AddPage(master.Id);
        document.AddPage(master.Id);
        Shape(0, 72, 144, 240, 120, OfficeColor.Red);
        PublisherPicture first = Picture(0, 108, null);
        Text(0, "Mixed first: overlap order");
        Shape(0, 72, 168, 240, 12, OfficeColor.Blue);
        Text(2, "Mixed last: overlap order");
        Picture(2, 108, first.ImageId);
        Shape(2, 72, 168, 240, 12, OfficeColor.FromRgb(0, 128, 0));
        Picture(2, 216, first.ImageId);
        return document;

        void Shape(int pageIndex, double x, double y, double width, double height, OfficeColor color) {
            var paint = OfficeShape.Rectangle(width, height);
            paint.FillColor = color; paint.StrokeColor = null;
            document.AddShape(document.Pages[pageIndex].Id, paint, OfficeLength.Points(x), OfficeLength.Points(y));
        }
        PublisherPicture Picture(int pageIndex, double x, int? imageId) {
            uint pageId = document.Pages[pageIndex].Id;
            PublisherPicture picture = imageId.HasValue
                ? document.AddImage(pageId, imageId.Value, OfficeLength.Points(x), OfficeLength.Points(144), OfficeLength.Points(96), OfficeLength.Points(96))
                : document.AddImage(pageId, image, OfficeLength.Points(x), OfficeLength.Points(144), OfficeLength.Points(96), OfficeLength.Points(96));
            return picture;
        }
        void Text(int pageIndex, string text) {
            var paragraph = new OfficeRichTextParagraph(new[] {
                new OfficeRichTextRun(text, 18, OfficeColor.Black, fontFamily: "Arial")
            });
            document.AddTextFrame(document.Pages[pageIndex].Id, OfficeLength.Points(72), OfficeLength.Points(144), OfficeLength.Points(288), OfficeLength.Points(108),
                new[] { paragraph });
        }
    }
}
