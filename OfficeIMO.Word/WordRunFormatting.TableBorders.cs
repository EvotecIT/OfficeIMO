using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal static partial class WordRunFormatting {
    /// <summary>Resolves a physical cell edge from authored cell, conditional, and table borders without changing source XML.</summary>
    internal static BorderType? ResolveTableCellBorder(WordTableCell cell, WordTableBorderSide side) =>
        ResolveTableCellBorder(cell, side, null);

    /// <summary>Resolves a diagonal cell border through the same source cascade as physical edges.</summary>
    internal static BorderType? ResolveTableCellDiagonalBorder(WordTableCell cell, bool descending) =>
        ResolveTableCellBorder(cell, null, descending);

    private static BorderType? ResolveTableCellBorder(WordTableCell cell, WordTableBorderSide? side, bool? descending) {
        Table? table = cell._tableCell.Ancestors<Table>().FirstOrDefault();
        bool rightToLeft = WordTableAlignmentResolver.IsRightToLeft(table?.GetFirstChild<TableProperties>());
        BorderType? direct = ReadBorder(cell._tableCell.TableCellProperties?.TableCellBorders);
        if (direct != null || table == null) return direct;
        TableRunPaintState state = GetTableRunPaintState(table, cell.Document.MainDocumentPartRoot.StyleDefinitionsPart?.Styles);
        if (state.HasConditionalFormatting) {
            foreach (TableStyleOverrideValues condition in EnumerateTableRunConditions(state, cell._tableCell)) {
                foreach (Style style in state.Chain) {
                    foreach (TableStyleProperties conditional in style.Elements<TableStyleProperties>()
                                 .Where(properties => properties.Type?.Value == condition).Reverse()) {
                        BorderType? border = ReadBorder(conditional.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>()?.GetFirstChild<TableCellBorders>());
                        if (border != null) return border;
                    }
                }
            }
        }
        foreach (Style style in state.Chain) {
            BorderType? border = ReadBorder(style.GetFirstChild<StyleTableCellProperties>()?.GetFirstChild<TableCellBorders>());
            if (border != null) return border;
        }
        if (!side.HasValue || !state.Cells.TryGetValue(cell._tableCell, out var position)) return null;
        int physicalColumn = rightToLeft ? state.ColumnCount - position.Column - position.Span : position.Column;
        BorderType? tableBorder = ReadTableBorder(table.GetFirstChild<TableProperties>()?.TableBorders);
        if (tableBorder != null) return tableBorder;
        foreach (Style style in state.Chain) {
            tableBorder = ReadTableBorder(style.StyleTableProperties?.TableBorders);
            if (tableBorder != null) return tableBorder;
        }
        return null;

        BorderType? ReadTableBorder(TableBorders? borders) => side switch {
            WordTableBorderSide.Top => position.Row == 0 ? borders?.TopBorder : borders?.InsideHorizontalBorder,
            WordTableBorderSide.Bottom => position.Row + cell.RowSpan >= state.RowCount ? borders?.BottomBorder : borders?.InsideHorizontalBorder,
            WordTableBorderSide.Left => physicalColumn == 0 ? ReadCellBorder(borders, side.Value, rightToLeft) : borders?.InsideVerticalBorder,
            WordTableBorderSide.Right => physicalColumn + position.Span >= state.ColumnCount ? ReadCellBorder(borders, side.Value, rightToLeft) : borders?.InsideVerticalBorder,
            _ => null
        };

        BorderType? ReadBorder(OpenXmlCompositeElement? borders) => side.HasValue
            ? ReadCellBorder(borders, side.Value, rightToLeft)
            : descending == true ? borders?.GetFirstChild<TopLeftToBottomRightCellBorder>()
                : (BorderType?)borders?.GetFirstChild<TopRightToBottomLeftCellBorder>();
    }

    private static BorderType? ReadCellBorder(OpenXmlCompositeElement? borders, WordTableBorderSide side, bool rightToLeft) => side switch {
        WordTableBorderSide.Top => borders?.GetFirstChild<TopBorder>(),
        WordTableBorderSide.Bottom => borders?.GetFirstChild<BottomBorder>(),
        WordTableBorderSide.Left => borders?.GetFirstChild<LeftBorder>() ??
            (rightToLeft ? (BorderType?)borders?.GetFirstChild<EndBorder>() : borders?.GetFirstChild<StartBorder>()),
        WordTableBorderSide.Right => borders?.GetFirstChild<RightBorder>() ??
            (rightToLeft ? (BorderType?)borders?.GetFirstChild<StartBorder>() : borders?.GetFirstChild<EndBorder>()),
        _ => null
    };
}
