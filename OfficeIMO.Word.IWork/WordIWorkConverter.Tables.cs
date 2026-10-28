using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

public static partial class WordIWorkConverter {
    private static WordTable? AddTable(WordDocument document, IWorkTable source,
        IWorkNativeListCatalog nativeLists,
        WordParagraph? pageHost, WordTable? tableHost, List<WordCellComment> cellComments, CancellationToken cancellationToken) {
        if (source.RowCount == 0 || source.ColumnCount == 0) return null;
        WordTable table = pageHost == null
            ? document.AddTable(source.RowCount, source.ColumnCount, WordTableStyle.TableGrid)
            : document.CreateTable(source.RowCount, source.ColumnCount, WordTableStyle.TableGrid);
        // Rows and Cells enumerate and rebuild wrapper lists on every access.
        // Keep row wrappers once and materialize cells only when a row needs them.
        List<WordTableRow> rows = table.Rows;
        var mergedBorders = new Dictionary<(int Row, int Column), OfficeBorderBox>();
        var hiddenEdge = new OfficeBorderSide(default, 0);
        foreach (var merge in source.ResolveMergedBorders(cancellationToken)) {
            for (int row = merge.Range.FirstRow; row <= merge.Range.LastRow; row++) {
                cancellationToken.ThrowIfCancellationRequested();
                // Word retains one horizontal-span cell per row of a vertical merge.
                mergedBorders[(row, merge.Range.FirstColumn)] = new OfficeBorderBox(merge.Borders.Left,
                    row == merge.Range.FirstRow ? merge.Borders.Top : hiddenEdge, merge.Borders.Right,
                    row == merge.Range.LastRow ? merge.Borders.Bottom : hiddenEdge);
            }
        }
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
        int?[] rowHeights = WordTableRowHeights(source, cancellationToken);
        bool hasBorderDeclarations = source.HasBorderDeclarations;

        for (int row = 1; row <= source.RowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (rowHeights[row - 1] is int height) {
                if (source.AutoResizeRows == true) rows[row - 1].MinimumHeight = global::OfficeIMO.OfficeLength.Points((height) / 20D);
                else rows[row - 1].Height = global::OfficeIMO.OfficeLength.Points((height) / 20D);
            }
        }
        for (int row = 1; row <= source.RowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            List<WordTableCell>? targetCells = null;
            for (int column = 1; column <= source.ColumnCount; column++) {
                var fill = source.GetFill(row, column);
                var borders = mergedBorders.TryGetValue((row, column), out var perimeter)
                    ? perimeter : source.GetBorders(row, column);
                IWorkParagraphStyle? style = source.GetParagraphStyle(row, column);
                IWorkCellPadding? padding = source.GetPadding(row, column);
                IWorkCellVerticalAlignment? vertical = source.GetVerticalAlignment(row, column);
                if (fill == null && style == null && padding == null && vertical == null && borders.Left == null && borders.Top == null && borders.Right == null && borders.Bottom == null && !hasBorderDeclarations) continue;
                targetCells ??= rows[row - 1].Cells;
                WordTableCell target = targetCells[column - 1];
                ApplyTableBorders(target, borders, hasBorderDeclarations);
                if (padding != null) {
                    target.MarginLeft = global::OfficeIMO.OfficeLength.Points((checked((short)ToSignedTwips(padding.LeftPoints))) / 20D);
                    target.MarginTop = global::OfficeIMO.OfficeLength.Points((checked((short)ToSignedTwips(padding.TopPoints))) / 20D);
                    target.MarginRight = global::OfficeIMO.OfficeLength.Points((checked((short)ToSignedTwips(padding.RightPoints))) / 20D);
                    target.MarginBottom = global::OfficeIMO.OfficeLength.Points((checked((short)ToSignedTwips(padding.BottomPoints))) / 20D);
                }
                if (vertical.HasValue) target.VerticalAlignment = vertical.Value switch {
                    IWorkCellVerticalAlignment.Top => WordTableVerticalAlignment.Top,
                    IWorkCellVerticalAlignment.Middle => WordTableVerticalAlignment.Center,
                    _ => WordTableVerticalAlignment.Bottom
                };


                if (fill != null) {
                    if (fill.Color is { } color) target.ShadingFillColorHex = color.RgbHex;
                    else target.ShadingPattern = WordShadingPattern.Nil;
                }
                if (style != null) {
                    WordParagraph paragraph = target.Paragraphs[0];
                    InitializeParagraphSpacing(paragraph);
                    ApplyParagraphStyle(paragraph, style, string.Empty);
                    ApplyParagraphTextStyle(paragraph, style.TextStyle);
                }
            }
        }
        int sourceRow = 0;
        List<WordTableCell>? sourceRowCells = null;
        foreach (IWorkTableCell sourceCell in source.Cells) {
            cancellationToken.ThrowIfCancellationRequested();
            if (sourceCell.Row != sourceRow) {
                sourceRow = sourceCell.Row;
                sourceRowCells = rows[sourceRow - 1].Cells;
            }
            WordTableCell target = sourceRowCells![sourceCell.Column - 1];
            IWorkParagraphStyle? defaultStyle = source.GetParagraphStyle(sourceCell.Row, sourceCell.Column);
            if (sourceCell.RichText is { Paragraphs.Count: > 0 } richText) {
                bool first = true;
                AddRichText(richText, _ => {
                    WordParagraph paragraph = target.AddParagraph(string.Empty,
                        removeExistingParagraphs: first);
                    first = false;
                    return paragraph;
                }, nativeLists, cancellationToken: cancellationToken, defaultStyle: defaultStyle);
            } else {
                sourceCell.TryGetFormattedNumber(out string cellText, out string? numericColor);
                WordParagraph paragraph = target.AddParagraph(cellText,
                    removeExistingParagraphs: true);
                InitializeParagraphSpacing(paragraph);
                if (defaultStyle != null) {
                    ApplyParagraphStyle(paragraph, defaultStyle, paragraph.Text);
                    ApplyParagraphTextStyle(paragraph, defaultStyle.TextStyle);
                }
                if (numericColor != null) paragraph.ColorHex = numericColor;
            }
            if (sourceCell.Comment is { } comment) {
                cellComments.Add(new WordCellComment(target, comment.Author, string.Empty,
                    comment.Text, comment.CreationDateUtc));
            }
        }
        foreach (IWorkTableMergeRange merge in source.MergedRanges) {
            cancellationToken.ThrowIfCancellationRequested();
            table.MergeCells(merge.FirstRow - 1, merge.FirstColumn - 1,
                merge.LastRow - merge.FirstRow + 1, merge.LastColumn - merge.FirstColumn + 1);
        }
        for (int row = 0; row < Math.Min(source.HeaderRowCount, rows.Count); row++) {
            rows[row].RepeatHeaderRowAtTheTopOfEachPage = true;
        }
        if (tableHost != null) tableHost._table.InsertAfterSelf(table._table);
        else if (pageHost != null) document.InsertTableAfter(pageHost, table);
        return table;
    }
}
