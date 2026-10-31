using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects the public table model, including its single complete story and four cell boundaries.</summary>
internal static class CandidateTableCreation {
    internal static CandidatePublicationSpec Create(bool edited) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        uint pageId = document.Pages[0].Id;
        PublisherTextFrame label = document.AddTextFrame(pageId, 72, 72, 300, 54, Paragraphs("Independent table label"));
        PublisherTable table = document.AddTable(pageId, 72, 144, new[] { 120D, 180D }, new[] { 42D, 54D });
        string[] values = { edited ? "Updated\nItem" : "Item", "Count", "Alpha", edited ? "7" : "" };
        for (int row = 0; row < 2; row++)
            for (int column = 0; column < 2; column++)
                document.ReplaceTableCellParagraphs(table.Id, row, column, Paragraphs(values[row * 2 + column], row == 0));
        table = document.Pages[0].Tables.Single();
        var objects = new[] {
            new CandidateObjectSpec(CandidateObjectKind.TextBox, new CandidatePlacement(label.X, label.Y, label.Width, label.Height),
                CandidateStorySpec.Styled(document.TextStories.Single(story => story.Id == label.StoryId).Paragraphs.Single()),
                ColumnSpacing: label.ColumnSpacing),
            new CandidateObjectSpec(CandidateObjectKind.Table, new CandidatePlacement(table.X, table.Y, table.Width, table.Height),
                CandidateStorySpec.Table(table, document.TextStories.Single(story => story.Id == table.StoryId)), Table: table)
        };
        return new CandidatePublicationSpec(612 * 12700, 792 * 12700, objects);
    }

    private static OfficeRichTextParagraph[] Paragraphs(string text, bool bold = false) => text.Split('\n').Select(value =>
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(value, 12, OfficeColor.Black, bold, fontFamily: "Arial") }, lineHeight: 18)).ToArray();
}
