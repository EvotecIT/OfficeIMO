using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Authors one newsletter combining the qualified native object, text-layout and table profiles.</summary>
internal static class CandidateNewsletter {
    internal static PublisherNativeCreationPlan Create(byte[] image) => PublisherNativeCreationProjection.ComposedContent(Author(image));

    internal static PublisherDocument Author(byte[] image) {
        PublisherDocument document = PublisherDocument.Create(OfficeLength.Points(612), OfficeLength.Points(792));
        PublisherPage master = document.AddMasterPage("Newsletter");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        document.AddPage(master.Id);
        document.AddPage(master.Id);
        Text(master.Id, "Community newsletter", 36, 30, 432, 42, 20, true);
        PublisherPicture logo = document.AddImage(master.Id, image, OfficeLength.Points(504), OfficeLength.Points(24), OfficeLength.Points(72), OfficeLength.Points(48));
        var rule = OfficeShape.Rectangle(540, 6);
        rule.FillColor = OfficeColor.Blue; rule.StrokeColor = null;
        document.AddShape(master.Id, rule, OfficeLength.Points(36), OfficeLength.Points(84));

        string body = string.Concat(Enumerable.Repeat(
            "Our community shares practical ideas, local news and upcoming activities. ", 12)).TrimEnd();
        PublisherTextFrame lead = Text(document.Pages[0].Id, body, 36, 120, 330, 180);
        document.AddLinkedTextFrame(lead.Id, document.Pages[2].Id, OfficeLength.Points(36), OfficeLength.Points(120), OfficeLength.Points(540), OfficeLength.Points(180));
        document.SetStoryAutomaticHyphenation(lead.StoryId, false);
        Text(document.Pages[0].Id, "In this issue: shared spaces, community projects and the events calendar.",
            390, 120, 186, 156);
        PublisherTable calendar = document.AddTable(document.Pages[0].Id, OfficeLength.Points(36), OfficeLength.Points(360),
            new OfficeLength[] { OfficeLength.Points(180D), OfficeLength.Points(150D) }, new OfficeLength[] { OfficeLength.Points(30D), OfficeLength.Points(42D) });
        string[] cells = { "Activity", "Date", "Community meeting", "Saturday" };
        for (int row = 0; row < 2; row++)
            for (int column = 0; column < 2; column++)
                document.ReplaceTableCellParagraphs(calendar.Id, row, column, Paragraphs(cells[row * 2 + column], 12, row == 0));

        PublisherTextFrame columns = Text(document.Pages[1].Id,
            string.Concat(Enumerable.Repeat("Small projects help neighbors share skills and improve shared spaces. ", 10)).TrimEnd(),
            36, 120, 540, 240);
        document.SetTextFrameLayout(columns.Id, 2, OfficeLength.Points(18), new OfficeTextPadding(6, 9, 12, 3), OfficeTextVerticalAlignment.Top);
        document.AddImage(document.Pages[1].Id, logo.ImageId!.Value, OfficeLength.Points(36), OfficeLength.Points(420), OfficeLength.Points(144), OfficeLength.Points(96));
        Text(document.Pages[2].Id, "Thank you for taking part. Send your ideas for the next issue.", 36, 360, 540, 72);
        return document;

        PublisherTextFrame Text(uint page, string text, double x, double y, double width, double height,
            double size = 12, bool bold = false) => document.AddTextFrame(page, OfficeLength.Points(x), OfficeLength.Points(y), OfficeLength.Points(width), OfficeLength.Points(height), Paragraphs(text, size, bold));
    }

    private static OfficeRichTextParagraph[] Paragraphs(string text, double size, bool bold) => new[] {
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black, bold, fontFamily: "Arial") }, lineHeight: size + 4)
    };
}
