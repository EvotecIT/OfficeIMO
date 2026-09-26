using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordTable {
    /// <summary>
    /// Replaces the content of the first paragraph in an indexed cell with a single text run.
    /// This is the allocation-efficient path for filling rectangular reports; use
    /// <see cref="WordTableRow.Cells"/> when richer cell content is required.
    /// </summary>
    /// <param name="rowIndex">Zero-based row index.</param>
    /// <param name="columnIndex">Zero-based cell index within the row.</param>
    /// <param name="text">Text to place in the cell.</param>
    /// <param name="bold">Whether the new run is bold.</param>
    /// <returns>The current table for fluent use.</returns>
    public WordTable SetCellText(int rowIndex, int columnIndex, string? text, bool bold = false) {
        if (columnIndex < 0) throw new ArgumentOutOfRangeException(nameof(columnIndex));
        List<TableRow> rows = GetRowElements();
        if ((uint)rowIndex >= (uint)rows.Count) throw new ArgumentOutOfRangeException(nameof(rowIndex));
        TableCell? cell = rows[rowIndex].GetFirstChild<TableCell>();
        for (int index = 0; index < columnIndex && cell != null; index++) {
            cell = cell.NextSibling<TableCell>();
        }
        if (cell == null) throw new ArgumentOutOfRangeException(nameof(columnIndex));

        Paragraph paragraph = cell.GetFirstChild<Paragraph>() ?? cell.AppendChild(new Paragraph());
        // Generated cells already contain an empty run. Fill it directly before taking the
        // general replacement path, which must still clear existing rich cell content.
        if (paragraph.FirstChild is ParagraphProperties properties &&
            properties.NextSibling() is Run emptyRun &&
            emptyRun.FirstChild == null &&
            emptyRun.NextSibling() == null) {
            if (bold) emptyRun.RunProperties = new RunProperties(new Bold());
            emptyRun.Append(new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
            return this;
        }

        OpenXmlElement? child = paragraph.FirstChild;
        while (child != null) {
            OpenXmlElement? next = child.NextSibling();
            if (child is not ParagraphProperties) child.Remove();
            child = next;
        }
        var run = new Run();
        if (bold) run.RunProperties = new RunProperties(new Bold());
        run.Append(new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
        paragraph.Append(run);
        return this;
    }
}
