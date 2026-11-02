using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects the public table model, including its complete story and rectangular or merged cell boundaries.</summary>
internal static class CandidateTableCreation {
    internal static CandidatePublicationSpec Create(bool edited) => CandidateModelProjection.Tables(Author(edited));

    internal static CandidatePublicationSpec CreateMerged() {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        uint page = document.Pages[0].Id;
        document.SetPageMaster(page, master.Id);
        document.AddTextFrame(page, 72, 72, 300, 54, Paragraphs("Independent table label"));
        PublisherTable table = document.AddTable(page, 72, 144, new[] { 90D, 110D, 130D, 150D }, new[] { 42D, 54D, 66D });
        document.MergeTableCells(table.Id, 0, 0, 1, 2);
        document.MergeTableCells(table.Id, 1, 0, 2, 1);
        document.MergeTableCells(table.Id, 1, 2, 2, 2);
        foreach (PublisherTableCell cell in document.Pages[0].Tables.Single().Cells) {
            string text = (cell.RowIndex, cell.ColumnIndex) switch {
                (0, 0) => "Horizontal", (1, 0) => "Vertical", (1, 2) => "Block\nSecond line",
                _ => $"R{cell.RowIndex + 1}C{cell.ColumnIndex + 1}"
            };
            document.ReplaceTableCellParagraphs(table.Id, cell.RowIndex, cell.ColumnIndex, Paragraphs(text, cell.RowIndex == 0));
        }
        return CandidateModelProjection.Tables(document);
    }

    internal static PublisherDocument Author(bool edited) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        uint pageId = document.Pages[0].Id;
        document.AddTextFrame(pageId, 72, 72, 300, 54, Paragraphs("Independent table label"));
        PublisherTable table = document.AddTable(pageId, 72, 144, new[] { 120D, 180D }, new[] { 42D, 54D });
        string[] values = { edited ? "Updated\nItem" : "Item", "Count", "Alpha", edited ? "7" : "" };
        for (int row = 0; row < 2; row++)
            for (int column = 0; column < 2; column++)
                document.ReplaceTableCellParagraphs(table.Id, row, column, Paragraphs(values[row * 2 + column], row == 0));
        return document;
    }

    private static OfficeRichTextParagraph[] Paragraphs(string text, bool bold = false) => text.Split('\n').Select(value =>
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(value, 12, OfficeColor.Black, bold, fontFamily: "Arial") }, lineHeight: 18)).ToArray();
}
