namespace OfficeIMO.Word;

public partial class WordDocument {
    private static void ApplyTableAlignments(WordTable table, OfficeTableOptions? options) {
        OfficeIMO.Core.Internal.ObjectDataHelpers.ValidateTableAlignments(options);
        if (options == null) return;
        foreach (var row in table.Rows) {
            for (int column = 0; column < row.Cells.Count && column < options.Alignments.Count; column++) {
                if (options.Alignments[column] is not { } alignment) continue;
                var native = alignment switch {
                    OfficeAlignment.Left => WordParagraphAlignment.Left,
                    OfficeAlignment.Start => WordParagraphAlignment.Start,
                    OfficeAlignment.Center => WordParagraphAlignment.Center,
                    OfficeAlignment.Right => WordParagraphAlignment.Right,
                    OfficeAlignment.End => WordParagraphAlignment.End,
                    _ => WordParagraphAlignment.Both
                };
                foreach (var paragraph in row.Cells[column].Paragraphs) paragraph.ParagraphAlignment = native;
            }
        }
    }
}
