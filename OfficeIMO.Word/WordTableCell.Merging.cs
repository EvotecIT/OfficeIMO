using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordTableCell {
    /// <summary>Merges this cell with the specified number of following logical cells.</summary>
    /// <param name="cellsCount">Number of additional cells to merge after this cell.</param>
    /// <param name="copyParagraphs">Moves merged cells' paragraphs into this cell when true; otherwise clears them.</param>
    public void MergeHorizontally(int cellsCount, bool copyParagraphs = false) {
        HorizontalMerge = WordCellMerge.Restart;
        foreach (TableCell next in WordTableContent.CellsAfter(_tableCell, cellsCount)) {
            MergeParagraphs(next, copyParagraphs);
            TableCellProperties properties = GetOrCreateTableCellProperties(next);
            properties.HorizontalMerge = new HorizontalMerge { Val = MergedCellValues.Continue };
            NormalizeTableCellPropertiesOrder(properties);
        }
    }

    /// <summary>Removes horizontal merge declarations from this cell and the specified following logical cells.</summary>
    /// <param name="cellsCount">Number of additional cells to split after this cell.</param>
    public void SplitHorizontally(int cellsCount) {
        HorizontalMerge = null;
        foreach (TableCell next in WordTableContent.CellsAfter(_tableCell, cellsCount))
            GetOrCreateTableCellProperties(next).HorizontalMerge?.Remove();
    }

    /// <summary>Merges this cell with cells at the same logical index in the specified following rows.</summary>
    /// <param name="cellsCount">Number of additional rows to merge below this cell.</param>
    /// <param name="copyParagraphs">Moves merged cells' paragraphs into this cell when true; otherwise clears them.</param>
    public void MergeVertically(int cellsCount, bool copyParagraphs = false) {
        VerticalMerge = WordCellMerge.Restart;
        foreach (TableCell next in WordTableContent.CellsBelow(_tableCell, cellsCount)) {
            MergeParagraphs(next, copyParagraphs);
            TableCellProperties properties = GetOrCreateTableCellProperties(next);
            properties.VerticalMerge = new VerticalMerge { Val = MergedCellValues.Continue };
            NormalizeTableCellPropertiesOrder(properties);
        }
    }

    /// <summary>Removes vertical merge declarations from this cell and cells at the same logical index below it.</summary>
    /// <param name="cellsCount">Number of additional rows to split below this cell.</param>
    public void SplitVertically(int cellsCount) {
        VerticalMerge = null;
        foreach (TableCell next in WordTableContent.CellsBelow(_tableCell, cellsCount))
            GetOrCreateTableCellProperties(next).VerticalMerge?.Remove();
    }

    private void MergeParagraphs(TableCell source, bool copyParagraphs) {
        foreach (Paragraph paragraph in source.Elements<Paragraph>().ToList()) {
            paragraph.Remove();
            if (copyParagraphs) _tableCell.Append(paragraph);
        }
        source.Append(new Paragraph());
    }
}
