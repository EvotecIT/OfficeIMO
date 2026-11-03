using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Authors one populated master and page-owned foreground without duplicating the master objects.</summary>
internal static class CandidateMasterContent {
    internal static PublisherNativeCreationPlan Create(byte[] image) => PublisherNativeCreationProjection.SingleMasterContent(Author(image));

    internal static PublisherDocument Author(byte[] image) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        document.AddPage(master.Id);
        document.AddPage(master.Id);
        Shape(master.Id, 36, 72, 240, 96, OfficeColor.Red);
        PublisherPicture logo = document.AddPicture(master.Id, image, 72, 72, 96, 96);
        Text(master.Id, "Master: shared content", 36, 72, 18);
        Shape(document.Pages[0].Id, 36, 84, 240, 12, OfficeColor.Blue);
        document.AddPicture(document.Pages[2].Id, logo.ImageId!.Value, 180, 72, 96, 96);
        Shape(document.Pages[2].Id, 36, 108, 240, 12, OfficeColor.FromRgb(0, 128, 0));
        Text(document.Pages[2].Id, "Page foreground", 36, 180, 12);
        return document;

        void Shape(uint owner, double x, double y, double width, double height, OfficeColor color) {
            var paint = OfficeShape.Rectangle(width, height);
            paint.FillColor = color; paint.StrokeColor = null;
            document.AddShape(owner, paint, x, y);
        }
        void Text(uint owner, string text, double x, double y, double size) {
            document.AddTextFrame(owner, x, y, 288, 108, new[] {
                new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black, fontFamily: "Arial") })
            });
        }
    }
}
