using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Composes the existing centered/paired master control with ordinary linked text, layout and physical-page content.</summary>
internal static class CandidatePairedBrochure {
    internal static PublisherNativeCreationPlan Create(byte[] image) => PublisherNativeCreationProjection.ComposedContent(Author(image));

    internal static PublisherDocument Author(byte[] image) {
        PublisherDocument document = CandidateMasterPages.Author(paired: true);
        PublisherPicture? logo = null;
        OfficeColor[] colors = { OfficeColor.Blue, OfficeColor.Red, OfficeColor.FromRgb(0, 128, 0) };
        for (int index = 0; index < document.MasterPages.Count; index++) {
            PublisherPage master = document.MasterPages[index];
            Rectangle(master.Id, 36, 66, 540, 6, colors[index]);
            double logoX = index == 1 ? 72 : 504;
            if (logo == null) logo = document.AddPicture(master.Id, image, logoX, 708, 72, 48);
            else document.AddPicture(master.Id, logo.ImageId!.Value, logoX, 708, 72, 48);
        }

        PublisherTextFrame lead = document.AddTextFrame(document.Pages[0].Id, 36, 180, 330, 180,
            CandidateMultiParagraph.Paragraphs());
        document.SetStoryAutomaticHyphenation(lead.StoryId, true);
        PublisherTextFrame middle = document.AddLinkedTextFrame(lead.Id, document.Pages[2].Id, 36, 180, 540, 180);
        document.AddLinkedTextFrame(middle.Id, document.Pages[1].Id, 36, 180, 330, 180);
        Text(document.Pages[0].Id, "In this brochure: shared spaces, community projects and the events calendar.",
            390, 180, 186, 156);
        Rectangle(document.Pages[0].Id, 36, 384, 540, 6, OfficeColor.Blue);
        PublisherTable calendar = document.AddTable(document.Pages[0].Id, 36, 420,
            new[] { 180D, 150D }, new[] { 30D, 42D });
        string[] cells = { "Activity", "Date", "Community meeting", "Saturday" };
        for (int row = 0; row < 2; row++)
            for (int column = 0; column < 2; column++)
                document.ReplaceTableCellParagraphs(calendar.Id, row, column, Paragraphs(cells[row * 2 + column], row == 0));

        PublisherTextFrame columns = Text(document.Pages[1].Id,
            string.Concat(Enumerable.Repeat("Small projects help neighbors share skills and improve shared spaces. ", 10)).TrimEnd(),
            36, 420, 540, 240);
        document.SetTextFrameLayout(columns.Id, 2, 18, new OfficeTextPadding(6, 9, 12, 3), OfficeTextVerticalAlignment.Top);
        document.AddPicture(document.Pages[1].Id, logo!.ImageId!.Value, 390, 180, 144, 96);
        Text(document.Pages[2].Id, "Thank you for taking part. Send your ideas for the next issue.", 36, 420, 540, 72);
        return document;

        void Rectangle(uint owner, double x, double y, double width, double height, OfficeColor color) {
            var paint = OfficeShape.Rectangle(width, height);
            paint.FillColor = color;
            paint.StrokeColor = null;
            document.AddShape(owner, paint, x, y);
        }

        PublisherTextFrame Text(uint owner, string text, double x, double y, double width, double height) {
            PublisherTextFrame frame = document.AddTextFrame(owner, x, y, width, height, Paragraphs(text));
            document.SetStoryAutomaticHyphenation(frame.StoryId, true);
            return frame;
        }
    }

    private static OfficeRichTextParagraph[] Paragraphs(string text, bool bold = false) => new[] {
        new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun(text, 12, OfficeColor.Black, bold, fontFamily: "Arial").WithLanguageId(1033)
        }, lineHeight: 16)
    };
}
