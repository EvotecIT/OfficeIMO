using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal static partial class WordRunFormatting {
    /// <summary>Resolves direct and inherited table-level shading, retaining explicit no-fill and excluding cell-style shading.</summary>
    internal static Shading? ResolveTableShading(WordTable table) {
        TableProperties? properties = table._table.GetFirstChild<TableProperties>();
        if (properties?.GetFirstChild<Shading>() is { } direct) return direct;
        Styles? styles = table.Document.MainDocumentPartRoot.StyleDefinitionsPart?.Styles;
        foreach (Style style in EnumerateStyleChain(styles, properties?.TableStyle?.Val?.Value, StyleValues.Table)) {
            if (style.StyleTableProperties?.GetFirstChild<Shading>() is { } shading) return shading;
        }
        return null;
    }

    /// <summary>Resolves cell shading in source override order, retaining explicit nil and clear declarations.</summary>
    internal static Shading? ResolveTableCellShading(WordTableCell cell) {
        if (cell._tableCell.TableCellProperties?.GetFirstChild<Shading>() is { } direct) return direct;
        Table? table = cell._tableCell.Ancestors<Table>().FirstOrDefault();
        if (table == null) return null;
        if (table.GetFirstChild<TableProperties>()?.GetFirstChild<Shading>() is { } tableShading) return tableShading;
        Styles? styles = cell.Document.MainDocumentPartRoot.StyleDefinitionsPart?.Styles;
        TableRunPaintState state = GetTableRunPaintState(table, styles);
        if (state.HasConditionalPaint) {
            foreach (TableStyleOverrideValues condition in EnumerateTableRunConditions(state, cell._tableCell)) {
                foreach (Style style in state.Chain) {
                    foreach (TableStyleProperties conditional in style.Elements<TableStyleProperties>()
                                 .Where(properties => properties.Type?.Value == condition).Reverse()) {
                        if (conditional.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>()?.GetFirstChild<Shading>() is { } shading)
                            return shading;
                    }
                }
            }
        }
        foreach (Style style in state.Chain) {
            Shading? shading = style.GetFirstChild<StyleTableCellProperties>()?.GetFirstChild<Shading>()
                ?? style.StyleTableProperties?.GetFirstChild<Shading>();
            if (shading != null) return shading;
        }
        return null;
    }

    /// <summary>
    /// Enumerates the containing table's conditional and whole-table run properties in override order.
    /// Call after direct/character/paragraph properties and before document defaults. Elements are
    /// returned unchanged so explicit highlight none and shading nil/clear still stop inheritance.
    /// Reads the nearest table's XML without constructing normalization-capable table wrappers.
    /// </summary>
    internal static IEnumerable<OpenXmlCompositeElement> EnumerateTableRunPropertySources(WordParagraph paragraph) {
        TableCell? cell = paragraph._paragraph.Ancestors<TableCell>().FirstOrDefault();
        TableRow? row = cell?.Ancestors<TableRow>().FirstOrDefault();
        Table? table = row?.Ancestors<Table>().FirstOrDefault();
        if (cell == null || row == null || table == null) yield break;

        Styles? styles = paragraph._document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        TableRunPaintState state = GetTableRunPaintState(table, styles);
        Style[] chain = state.Chain;
        if (chain.Length == 0) yield break;

        if (state.HasConditionalPaint) {
            foreach (TableStyleOverrideValues condition in EnumerateTableRunConditions(state, cell)) {
                foreach (Style style in chain) {
                    foreach (TableStyleProperties conditional in style.Elements<TableStyleProperties>()
                                 .Where(properties => properties.Type?.Value == condition).Reverse()) {
                        RunPropertiesBaseStyle? properties = conditional.GetFirstChild<RunPropertiesBaseStyle>();
                        if (HasRunBackground(properties)) yield return properties!;
                    }
                }
            }
        }

        foreach (Style style in chain) {
            if (HasRunBackground(style.StyleRunProperties)) yield return style.StyleRunProperties!;
        }
    }

    private static bool HasRunBackground(OpenXmlCompositeElement? properties) =>
        properties?.GetFirstChild<Shading>() != null || ReadRunHighlight(properties).HasValue;

    private static bool IsBuiltInNormalTableStyle(string? styleId) =>
        string.Equals(styleId, "TableNormal", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(styleId, "NormalTable", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Resolves band-one and band-two run paint before the edge conditions, including the
    /// contiguous repeated-header/body/footer boundaries. Logical columns include omitted
    /// leading cells and merged spans; corner overrides remain outside this supported set.
    /// </summary>
    private static IEnumerable<TableStyleOverrideValues> EnumerateTableRunConditions(TableRunPaintState state, TableCell cell) {
        if (!state.Cells.TryGetValue(cell, out var position)) yield break;
        int rowIndex = position.Row, column = position.Column, mask = state.Mask;
        int headerRows = state.HeaderRows, footerStart = state.FooterStart, columnCount = state.ColumnCount;
        bool firstRow = (mask & 0x0020) != 0, lastRow = (mask & 0x0040) != 0;

        // Reverse the native low-to-high merge order. An inherited condition outranks
        // a derived whole-table property, while a derived property wins within a condition.
        if (lastRow && rowIndex >= footerStart)
            yield return TableStyleOverrideValues.LastRow;
        if (firstRow && rowIndex == 0)
            yield return TableStyleOverrideValues.FirstRow;
        if ((mask & 0x0100) != 0 && columnCount > 0 && (long)column + position.Span >= columnCount)
            yield return TableStyleOverrideValues.LastColumn;
        if ((mask & 0x0080) != 0 && column == 0)
            yield return TableStyleOverrideValues.FirstColumn;
        if (rowIndex >= headerRows && rowIndex < footerStart) {
            int bandColumn = column - ((mask & 0x0080) != 0 ? 1 : 0);
            if ((mask & 0x0400) == 0 && state.ColumnBandSize > 0 && bandColumn >= 0)
                yield return bandColumn / state.ColumnBandSize % 2 == 0
                    ? TableStyleOverrideValues.Band1Vertical : TableStyleOverrideValues.Band2Vertical;
            if ((mask & 0x0200) == 0 && state.RowBandSize > 0)
                yield return (rowIndex - headerRows) / state.RowBandSize % 2 == 0
                    ? TableStyleOverrideValues.Band1Horizontal : TableStyleOverrideValues.Band2Horizontal;
        }
    }
}
