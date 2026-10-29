using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Applies the theme's numeric alignment when assigning a plain, otherwise unaligned table value.</summary>
    private void ApplyDefaultTableNumberAlignment() {
        if (_paragraph.ParagraphProperties?.Justification != null ||
            _paragraph.ParagraphProperties?.ParagraphStyleId != null ||
            _paragraph.Elements<Run>().Count() != 1) return;
        TableCell? cell = _paragraph.Parent as TableCell;
        if (cell == null || cell.Elements<Paragraph>().Count() != 1 || cell.Elements<Table>().Any()) return;
        Table? table = cell.Ancestors<Table>().FirstOrDefault();
        if (table?.GetFirstChild<TableProperties>()?.TableStyle?.Val?.Value != "OfficeIMOModern" ||
            !WordNumericText.IsNumber(Text)) return;
        ParagraphAlignment = _document.EffectiveCreationTheme.TableStyle.NumberAlignment switch {
            OfficeAlignment.Right => WordParagraphAlignment.Right,
            OfficeAlignment.Center => WordParagraphAlignment.Center,
            OfficeAlignment.Start => WordParagraphAlignment.Start,
            OfficeAlignment.End => WordParagraphAlignment.End,
            _ => WordParagraphAlignment.Left
        };
    }
}
