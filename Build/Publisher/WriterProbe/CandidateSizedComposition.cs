using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Authors the same mixed-content publication at dimensions supplied as data in native EMU.</summary>
internal static class CandidateSizedComposition {
    internal static PublisherNativeCreationPlan Create(string? dimensions, byte[] image) {
        string[] values = (dimensions ?? string.Empty).Split('x');
        if (values.Length != 2
            || !uint.TryParse(values[0], NumberStyles.None, CultureInfo.InvariantCulture, out uint width)
            || !uint.TryParse(values[1], NumberStyles.None, CultureInfo.InvariantCulture, out uint height)
            || width == 0 || height == 0)
            throw new ArgumentException("The sized composition requires positive WIDTH_EMUxHEIGHT_EMU dimensions.", nameof(dimensions));
        return PublisherNativeCreationProjection.ComposedContent(Author(width, height, image));
    }

    internal static PublisherDocument Author(uint widthEmu, uint heightEmu, byte[] image) {
        double w = widthEmu / 12700D, h = heightEmu / 12700D;
        PublisherDocument document = PublisherDocument.Create(w, h);
        PublisherPage first = document.Pages[0];
        PublisherPage master = document.AddMasterPage("Community update");
        document.SetPageMaster(first.Id, master.Id);
        PublisherPage second = document.AddPage(master.Id);

        Text(master.Id, "Community update", 18, 12, w - 36, 14, 8, 12);
        Rectangle(master.Id, 18, 30, w - 36, 3, OfficeColor.Blue);
        PublisherPicture logo = document.AddPicture(master.Id, image, w - 54, h - 42, 36, 24);
        Text(first.Id, "Events and projects", 18, 48, w - 36, 30, 16, 20, true);
        PublisherTextFrame lead = Text(first.Id,
            "Neighbors share practical ideas, local news and plans for our shared spaces. Join the next meeting to help choose activities for the coming season.",
            18, 90, w - 36, 14, 10, 14);
        document.AddLinkedTextFrame(lead.Id, second.Id, 18, 48, w - 36, 56);

        PublisherTable calendar = document.AddTable(first.Id, 18, h - 96,
            new[] { 84D, w - 120 }, new[] { 18D, 18D });
        string[] cells = { "Activity", "Date", "Meeting", "Saturday" };
        for (int row = 0; row < 2; row++)
            for (int column = 0; column < 2; column++)
                document.ReplaceTableCellParagraphs(calendar.Id, row, column,
                    Paragraphs(cells[row * 2 + column], 9, 12, row == 0));

        PublisherTextFrame columns = Text(second.Id,
            string.Concat(Enumerable.Repeat("Share your ideas. Meet your neighbors. Plan the next community project. ", 2)).TrimEnd(),
            18, 120, w - 36, 48, 9, 12);
        document.SetTextFrameLayout(columns.Id, 2, 12, new OfficeTextPadding(3, 3, 3, 3), OfficeTextVerticalAlignment.Top);
        document.AddPicture(first.Id, logo.ImageId!.Value, w - 72, 132, 54, 36);
        Rectangle(second.Id, w - 54, h - 54, 36, 12, OfficeColor.Red);
        return document;

        PublisherTextFrame Text(uint owner, string text, double x, double y, double width, double height,
            double size, double lineHeight, bool bold = false) {
            PublisherTextFrame frame = document.AddTextFrame(owner, x, y, width, height, Paragraphs(text, size, lineHeight, bold));
            document.SetStoryAutomaticHyphenation(frame.StoryId, true);
            return frame;
        }

        void Rectangle(uint owner, double x, double y, double width, double height, OfficeColor color) {
            var paint = OfficeShape.Rectangle(width, height);
            paint.FillColor = color;
            paint.StrokeColor = null;
            document.AddShape(owner, paint, x, y);
        }
    }

    private static OfficeRichTextParagraph[] Paragraphs(string text, double size, double lineHeight, bool bold = false) => new[] {
        new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun(text, size, OfficeColor.Black, bold, fontFamily: "Arial").WithLanguageId(1033)
        }, lineHeight: lineHeight)
    };
}
