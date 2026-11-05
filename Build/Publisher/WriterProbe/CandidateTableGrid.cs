using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Authors legible rectangular and empty grid specimens through the canonical table/story model.</summary>
internal static class CandidateTableGrid {
    internal static PublisherNativeCreationPlan Create(string? dimensions, bool empty = false) =>
        PublisherNativeCreationProjection.Tables(Author(dimensions, empty));

    internal static PublisherNativeCreationPlan CreateEdited(string? dimensions) {
        PublisherDocument document = Author(dimensions);
        PublisherPage page = document.Pages[0];
        PublisherTable table = page.Tables.Single();
        PublisherTextFrame caption = page.TextFrames.Single();
        document.SetTableBounds(table.Id, OfficeLength.Points(72), OfficeLength.Points(132), OfficeLength.Points(432), OfficeLength.Points(table.Height * 0.9));
        document.SetTextFrameBounds(caption.Id, OfficeLength.Points(72), OfficeLength.Points(48), OfficeLength.Points(432), OfficeLength.Points(42));
        document.ReplaceTableCellParagraphs(table.Id, table.RowCount - 1, table.ColumnCount - 1, Paragraphs("Final\nCell", table.RowCount == 1));
        document.SetPageObjectOrder(page.Id, new[] { table.Id, caption.Id });
        return PublisherNativeCreationProjection.Tables(document);
    }

    internal static PublisherDocument Author(string? dimensions, bool empty = false) {
        string[] parts = (dimensions ?? string.Empty).Split('x');
        if (parts.Length != 2 || !int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out int rows)
            || !int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out int columns)
            || rows is < 1 or > 12 || columns is < 1 or > 8)
            throw new ArgumentException("The grid specimen requires <rows>x<columns>, with 1..12 rows and 1..8 columns.");
        PublisherDocument document = PublisherDocument.Create(OfficeLength.Points(612), OfficeLength.Points(792));
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        uint page = document.Pages[0].Id;
        document.AddTextFrame(page, OfficeLength.Points(66), OfficeLength.Points(48), OfficeLength.Points(480), OfficeLength.Points(42), Paragraphs($"Borderless table {rows} x {columns}", true));
        PublisherTable table = document.AddTable(page, OfficeLength.Points(66), OfficeLength.Points(120),
            Enumerable.Repeat(OfficeLength.Points(480D / columns), columns).ToArray(),
            Enumerable.Range(0, rows).Select(row => 42D + row % 3 * 6).Select(OfficeLength.Points).ToArray());
        for (int row = 0; row < rows; row++) {
            for (int column = 0; column < columns; column++) {
                bool first = row == 0 && column == 0, last = row == rows - 1 && column == columns - 1;
                string value = empty || first && rows * columns > 1 ? string.Empty : $"R{row + 1}C{column + 1}";
                if (!empty && last) value += "\nEnd";
                document.ReplaceTableCellParagraphs(table.Id, row, column, Paragraphs(value, row == 0));
            }
        }
        return document;
    }

    private static OfficeRichTextParagraph[] Paragraphs(string text, bool bold) => text.Split('\n').Select(value =>
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(value, 10, OfficeColor.Black, bold, fontFamily: "Arial") }, lineHeight: 14)).ToArray();
}
