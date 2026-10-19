using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal static partial class WordRunFormatting {
    /// <summary>
    /// Enumerates the containing table's conditional and whole-table run properties in override order.
    /// Call after direct/character/paragraph properties and before document defaults. Elements are
    /// returned unchanged so explicit highlight none and shading nil/clear still stop inheritance.
    /// Reads the nearest table's XML without constructing normalization-capable table wrappers.
    /// </summary>
    internal static IEnumerable<OpenXmlCompositeElement> EnumerateTableRunPropertySources(WordParagraph paragraph) {
        TableCell? cell = paragraph._paragraph.Ancestors<TableCell>().FirstOrDefault();
        TableRow? row = cell?.Parent as TableRow;
        Table? table = row?.Parent as Table;
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
    /// Mirrors the native converter's supported condition set and priority, including its
    /// contiguous repeated-header/body/footer boundaries. Logical columns include omitted
    /// leading cells and merged spans. Corner and band-two overrides are outside that set.
    /// </summary>
    private static IEnumerable<TableStyleOverrideValues> EnumerateTableRunConditions(TableRunPaintState state, TableCell cell) {
        if (!state.Cells.TryGetValue(cell, out var position)) yield break;
        int rowIndex = position.Row, column = position.Column, mask = state.Mask;
        int headerRows = state.HeaderRows, footerStart = state.FooterStart, columnCount = state.ColumnCount;
        bool firstRow = (mask & 0x0020) != 0, lastRow = (mask & 0x0040) != 0;

        // Reverse the native low-to-high merge order. An inherited condition outranks
        // a derived whole-table property, while a derived property wins within a condition.
        if ((mask & 0x0100) != 0 && columnCount > 0 && (long)column + position.Span >= columnCount)
            yield return TableStyleOverrideValues.LastColumn;
        if ((mask & 0x0080) != 0 && column == 0)
            yield return TableStyleOverrideValues.FirstColumn;
        if (rowIndex >= headerRows && rowIndex < footerStart) {
            if ((mask & 0x0400) == 0 && column % 2 == 1)
                yield return TableStyleOverrideValues.Band1Vertical;
            if ((mask & 0x0200) == 0 && (rowIndex - headerRows) % 2 == 1)
                yield return TableStyleOverrideValues.Band1Horizontal;
        }
        if (lastRow && rowIndex >= footerStart)
            yield return TableStyleOverrideValues.LastRow;
        if (firstRow && rowIndex == 0)
            yield return TableStyleOverrideValues.FirstRow;
    }
}
