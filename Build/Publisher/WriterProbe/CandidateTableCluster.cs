using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Authors the dense comparison grids through the same table/story model as other specimens.</summary>
internal static class CandidateTableCluster {
    internal static CandidatePublicationSpec Create(int size, bool edited) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        uint page = document.Pages[0].Id;
        document.SetPageMaster(page, master.Id);
        document.AddTextFrame(page, 36, 60, 480, 42, Paragraphs($"Independent dense {size}x{size} table", 12, false, null));
        PublisherTable table = document.AddTable(page, 36, 144, Enumerable.Repeat(480D / size, size).ToArray(),
            Enumerable.Repeat(360D / size, size).ToArray());
        var side = new OfficeBorderSide(OfficeColor.FromRgb(0, 128, 0), 0.5);
        for (int row = 0; row < size; row++) {
            for (int column = 0; column < size; column++) {
                document.ReplaceTableCellParagraphs(table.Id, row, column, Paragraphs($"{row + 1:D2}{column + 1:D2}", 8, row == 0, 12));
                OfficeColor? fill = ((row + column) % 3) switch {
                    0 => OfficeColor.FromRgb(255, 224, 160), 1 => OfficeColor.FromRgb(160, 224, 255), _ => null
                };
                if (edited && row == 0 && column == 2) fill = OfficeColor.FromRgb(180, 240, 180);
                if (fill.HasValue) document.SetTableCellFillColor(table.Id, row, column, fill);
            }
        }
        for (int row = 0; row <= size; row++) document.SetTableBorderStyle(table.Id, row, 0, row, size, side);
        for (int column = 0; column <= size; column++) document.SetTableBorderStyle(table.Id, 0, column, size, column, side);
        return CandidateModelProjection.Tables(document);
    }

    private static OfficeRichTextParagraph[] Paragraphs(string text, double size, bool bold, double? lineHeight) => new[] {
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black, bold, fontFamily: "Arial") }, lineHeight: lineHeight)
    };
}
