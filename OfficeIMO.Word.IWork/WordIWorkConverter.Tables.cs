using System.Threading;
using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

public static partial class WordIWorkConverter {
    private static WordTable? AddTable(WordDocument document, IWorkTable source,
        IWorkNativeListCatalog nativeLists,
        WordParagraph? pageHost, WordTable? tableHost, CancellationToken cancellationToken) {
        if (source.RowCount == 0 || source.ColumnCount == 0) return null;
        WordTable table = pageHost == null
            ? document.AddTable(source.RowCount, source.ColumnCount, WordTableStyle.TableGrid)
            : document.CreateTable(source.RowCount, source.ColumnCount, WordTableStyle.TableGrid);
        table.Description = source.AccessibilityDescription;
        if (source.DefaultColumnWidth is > 0 || source.ColumnWidths.Count > 0) {
            List<int> widths = table.ColumnWidth;
            table.ColumnWidthType = WordTableWidthUnit.Dxa;
            for (int column = 1; column <= source.ColumnCount; column++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (source.GetColumnWidth(column) is double width) widths[column - 1] = ToSignedTwips(width);
            }
            table.ColumnWidth = widths;
        }
        for (int row = 1; row <= source.RowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (source.GetRowHeight(row) is double height) {
                if (source.AutoResizeRows == true) table.Rows[row - 1].MinimumHeight = ToSignedTwips(height);
                else table.Rows[row - 1].Height = ToSignedTwips(height);
            }
        }
        for (int row = 1; row <= source.RowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int column = 1; column <= source.ColumnCount; column++) {
                WordTableCell fillTarget = table.Rows[row - 1].Cells[column - 1];
                if (source.GetFill(row, column) is { } fill) {
                    if (fill.Color is { } color) fillTarget.ShadingFillColorHex = color.RgbHex;
                    else fillTarget.ShadingPattern = WordShadingPattern.Nil;
                }
                if (source.GetParagraphStyle(row, column) is { } style) {
                    WordParagraph paragraph = table.Rows[row - 1].Cells[column - 1].Paragraphs[0];
                    ApplyParagraphStyle(paragraph, style, string.Empty);
                    ApplyTextStyle(paragraph, style.TextStyle);
                }
            }
        }
        foreach (IWorkTableCell sourceCell in source.Cells) {
            cancellationToken.ThrowIfCancellationRequested();
            WordTableCell target = table.Rows[sourceCell.Row - 1].Cells[sourceCell.Column - 1];
            if (sourceCell.Padding is { } padding) {
                target.MarginLeftWidth = checked((short)ToSignedTwips(padding.LeftPoints));
                target.MarginTopWidth = checked((short)ToSignedTwips(padding.TopPoints));
                target.MarginRightWidth = checked((short)ToSignedTwips(padding.RightPoints));
                target.MarginBottomWidth = checked((short)ToSignedTwips(padding.BottomPoints));
            }
            if (sourceCell.VerticalAlignment is { } vertical) target.VerticalAlignment = vertical switch {
                IWorkCellVerticalAlignment.Top => WordTableVerticalAlignment.Top,
                IWorkCellVerticalAlignment.Middle => WordTableVerticalAlignment.Center,
                _ => WordTableVerticalAlignment.Bottom
            };
            IWorkParagraphStyle? defaultStyle = source.GetParagraphStyle(sourceCell.Row, sourceCell.Column);
            bool header = sourceCell.Row <= source.HeaderRowCount
                || sourceCell.Column <= source.HeaderColumnCount
                || sourceCell.Row > source.RowCount - source.FooterRowCount;
            if (sourceCell.RichText is { Paragraphs.Count: > 0 } richText) {
                bool first = true;
                AddRichText(richText, _ => {
                    WordParagraph paragraph = target.AddParagraph(string.Empty,
                        removeExistingParagraphs: first);
                    first = false;
                    return paragraph;
                }, nativeLists, forceBold: header && defaultStyle?.TextStyle.Bold == null, cancellationToken: cancellationToken, defaultStyle: defaultStyle);
            } else {
                sourceCell.TryGetFormattedNumber(out string cellText, out string? numericColor);
                WordParagraph paragraph = target.AddParagraph(cellText,
                    removeExistingParagraphs: true);
                if (header) paragraph.Bold = true;
                if (defaultStyle != null) {
                    ApplyParagraphStyle(paragraph, defaultStyle, paragraph.Text);
                    ApplyTextStyle(paragraph, defaultStyle.TextStyle);
                }
                if (numericColor != null) paragraph.ColorHex = numericColor;
            }
        }
        foreach (IWorkTableMergeRange merge in source.MergedRanges) {
            cancellationToken.ThrowIfCancellationRequested();
            table.MergeCells(merge.FirstRow - 1, merge.FirstColumn - 1,
                merge.LastRow - merge.FirstRow + 1, merge.LastColumn - merge.FirstColumn + 1);
        }
        for (int row = 0; row < Math.Min(source.HeaderRowCount, table.Rows.Count); row++) {
            table.Rows[row].RepeatHeaderRowAtTheTopOfEachPage = true;
        }
        if (tableHost != null) tableHost._table.InsertAfterSelf(table._table);
        else if (pageHost != null) document.InsertTableAfter(pageHost, table);
        return table;
    }
}
