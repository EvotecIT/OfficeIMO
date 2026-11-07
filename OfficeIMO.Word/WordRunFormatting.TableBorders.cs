using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal static partial class WordRunFormatting {
    /// <summary>Resolves a physical cell edge from authored cell, conditional, and table borders without changing source XML.</summary>
    internal static BorderType? ResolveTableCellBorder(WordTableCell cell, WordTableBorderSide side) {
        Table? table = cell._tableCell.Ancestors<Table>().FirstOrDefault();
        bool rightToLeft = WordTableAlignmentResolver.IsRightToLeft(table?.GetFirstChild<TableProperties>());
        BorderType? direct = ReadCellBorder(cell._tableCell.TableCellProperties?.TableCellBorders, side, rightToLeft);
        if (direct != null || table == null) return direct;
        TableRunPaintState state = GetTableRunPaintState(table, cell.Document.MainDocumentPartRoot.StyleDefinitionsPart?.Styles);
        if (state.HasConditionalPaint) {
            foreach (TableStyleOverrideValues condition in EnumerateTableRunConditions(state, cell._tableCell)) {
                foreach (Style style in state.Chain) {
                    foreach (TableStyleProperties conditional in style.Elements<TableStyleProperties>()
                                 .Where(properties => properties.Type?.Value == condition).Reverse()) {
                        BorderType? border = ReadCellBorder(conditional.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>()?.GetFirstChild<TableCellBorders>(), side, rightToLeft);
                        if (border != null) return border;
                    }
                }
            }
        }
        foreach (Style style in state.Chain) {
            BorderType? border = ReadCellBorder(style.GetFirstChild<StyleTableCellProperties>()?.GetFirstChild<TableCellBorders>(), side, rightToLeft);
            if (border != null) return border;
        }
        if (!state.Cells.TryGetValue(cell._tableCell, out var position)) return null;
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
            WordTableBorderSide.Left => physicalColumn == 0 ? ReadCellBorder(borders, side, rightToLeft) : borders?.InsideVerticalBorder,
            WordTableBorderSide.Right => physicalColumn + position.Span >= state.ColumnCount ? ReadCellBorder(borders, side, rightToLeft) : borders?.InsideVerticalBorder,
            _ => null
        };
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
