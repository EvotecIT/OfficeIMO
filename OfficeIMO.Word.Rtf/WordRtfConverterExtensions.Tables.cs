using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyWordElements(IEnumerable<WordElement> elements, RtfDocument destination, Dictionary<string, int> revisionAuthorIndexes) {
        var seenParagraphs = new HashSet<Paragraph>();
        foreach (WordElement element in elements) {
            if (element is WordParagraph wordParagraph) {
                if (!seenParagraphs.Add(wordParagraph._paragraph)) {
                    continue;
                }

                if (TryCopyImageBlock(wordParagraph, destination)) {
                    continue;
                }

                RtfParagraph paragraph = destination.AddParagraph();
                CopyTabStops(wordParagraph, paragraph);
                CopyParagraphFormatting(wordParagraph, paragraph, destination);
                AppendFormattedRuns(wordParagraph, paragraph, destination, revisionAuthorIndexes);
            } else if (element is WordTable table) {
                CopyTable(table, destination.AddTable(0, GetColumnCount(table)), destination, revisionAuthorIndexes, 1);
            }
        }
    }

    private static void CopyWordElements(IEnumerable<WordElement> elements, RtfSection destination, RtfDocument document, Dictionary<string, int> revisionAuthorIndexes) {
        var seenParagraphs = new HashSet<Paragraph>();
        foreach (WordElement element in elements) {
            if (element is WordParagraph wordParagraph) {
                if (!seenParagraphs.Add(wordParagraph._paragraph)) {
                    continue;
                }

                if (TryCopyImageBlock(wordParagraph, destination)) {
                    continue;
                }

                RtfParagraph paragraph = destination.AddParagraph();
                CopyTabStops(wordParagraph, paragraph);
                CopyParagraphFormatting(wordParagraph, paragraph, document);
                AppendFormattedRuns(wordParagraph, paragraph, document, revisionAuthorIndexes);
            } else if (element is WordTable table) {
                CopyTable(table, destination.AddTable(0, GetColumnCount(table)), document, revisionAuthorIndexes, 1);
            }
        }
    }

    private static void CopyTable(WordTable source, RtfTable destination, RtfDocument document, Dictionary<string, int> revisionAuthorIndexes, int tableDepth) {
        RtfTableTraversalGuard.EnsureDepth(tableDepth);
        int? cellGap = GetWordTableCellSpacing(source);
        int? tableLeftIndent = GetWordTableLeftIndent(source);
        RtfTableAlignment? tableAlignment = ToRtfTableAlignment(source.Alignment);
        RtfTableWidthUnit? tableWidthUnit = ToRtfTableWidthUnit(source.WidthType);
        int? tableWidth = source.Width;
        foreach (WordTableRow wordRow in source.Rows) {
            RtfTableRow row = destination.AddRow();
            row.RepeatHeader = wordRow.RepeatHeaderRowAtTheTopOfEachPage;
            row.KeepTogether = !wordRow.AllowRowToBreakAcrossPages;
            row.HeightTwips = wordRow.Height;
            row.CellGapTwips = cellGap;
            row.LeftIndentTwips = tableLeftIndent;
            row.Alignment = tableAlignment;
            row.PreferredWidthUnit = tableWidthUnit;
            row.PreferredWidth = tableWidth;
            int rightBoundary = 0;
            foreach (WordTableCell wordCell in wordRow.GetCells(readOnly: true)) {
                int width = wordCell.WidthType == TableWidthUnitValues.Dxa && wordCell.Width.HasValue
                    ? Math.Max(1, wordCell.Width.Value)
                    : 2400;
                rightBoundary += width;
                RtfTableCell cell = row.AddCell(rightBoundary);
                CopyCellProperties(wordCell, cell, document);
                CopyCellParagraphs(wordCell, cell, document, revisionAuthorIndexes, tableDepth);
            }
        }
    }

    private static void CopyCellProperties(WordTableCell source, RtfTableCell destination, RtfDocument document) {
        destination.HorizontalMerge = ToRtfCellMerge(source.HorizontalMerge);
        destination.VerticalMerge = ToRtfCellMerge(source.VerticalMerge);
        destination.VerticalAlignment = ToRtfCellVerticalAlignment(source.VerticalAlignment);
        destination.TextFlow = ToRtfCellTextFlow(source.TextDirection);
        destination.PreferredWidthUnit = ToRtfTableWidthUnit(source.WidthType);
        destination.PreferredWidth = source.Width;
        destination.NoWrap = !source.WrapText;
        destination.FitText = source.FitText;

        string shading = source.ShadingFillColorHex;
        if (!string.IsNullOrWhiteSpace(shading) && TryParseHexColor(shading, out byte red, out byte green, out byte blue)) {
            destination.BackgroundColorIndex = GetOrAddColor(document, red, green, blue);
        }

        string? foreground = source._tableCellProperties?.Shading?.Color?.Value;
        if (!string.IsNullOrWhiteSpace(foreground) &&
            !string.Equals(foreground, "auto", StringComparison.OrdinalIgnoreCase) &&
            TryParseHexColor(foreground!, out red, out green, out blue)) {
            destination.ShadingForegroundColorIndex = GetOrAddColor(document, red, green, blue);
        }

        CopyCellShadingPattern(source.ShadingPattern, destination);
        CopyCellPadding(source, destination);
        CopyCellBorders(source, destination, document);
    }

    private static void CopyCellShadingPattern(ShadingPatternValues? source, RtfTableCell destination) {
        if (!source.HasValue) {
            return;
        }

        int? shading = ToRtfShadingPercent(source.Value);
        if (shading.HasValue) {
            destination.ShadingPatternPercent = shading;
        }

        destination.ShadingPattern = ToRtfShadingPattern(source.Value);
    }

    private static void CopyCellPadding(WordTableCell source, RtfTableCell destination) {
        destination.PaddingTopTwips = source.MarginTopWidth;
        destination.PaddingLeftTwips = source.MarginLeftWidth;
        destination.PaddingBottomTwips = source.MarginBottomWidth;
        destination.PaddingRightTwips = source.MarginRightWidth;
    }

    private static void CopyCellBorders(WordTableCell source, RtfTableCell destination, RtfDocument document) {
        CopyCellBorder(source.Borders.TopStyle, source.Borders.TopSize?.Value, source.Borders.TopColorHex, destination.TopBorder, document);
        CopyCellBorder(source.Borders.LeftStyle, source.Borders.LeftSize?.Value, source.Borders.LeftColorHex, destination.LeftBorder, document);
        CopyCellBorder(source.Borders.BottomStyle, source.Borders.BottomSize?.Value, source.Borders.BottomColorHex, destination.BottomBorder, document);
        CopyCellBorder(source.Borders.RightStyle, source.Borders.RightSize?.Value, source.Borders.RightColorHex, destination.RightBorder, document);
        CopyCellBorder(source.Borders.TopLeftToBottomRightStyle, source.Borders.TopLeftToBottomRightSize?.Value, source.Borders.TopLeftToBottomRightColorHex, destination.TopLeftToBottomRightBorder, document);
        CopyCellBorder(source.Borders.TopRightToBottomLeftStyle, source.Borders.TopRightToBottomLeftSize?.Value, source.Borders.TopRightToBottomLeftColorHex, destination.TopRightToBottomLeftBorder, document);
    }

    private static void CopyCellBorder(BorderValues? style, uint? width, string? colorHex, RtfTableCellBorder destination, RtfDocument document) {
        destination.Style = ToRtfBorderStyle(style);
        if (width.HasValue) {
            destination.Width = checked((int)width.Value);
        }

        if (!string.IsNullOrWhiteSpace(colorHex) &&
            !string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase) &&
            TryParseHexColor(colorHex!, out byte red, out byte green, out byte blue)) {
            destination.ColorIndex = GetOrAddColor(document, red, green, blue);
        }
    }

    private static void CopyCellParagraphs(WordTableCell source, RtfTableCell destination, RtfDocument document, Dictionary<string, int> revisionAuthorIndexes, int tableDepth) {
        foreach (OpenXmlElement child in source._tableCell.ChildElements) {
            if (child is Paragraph wordParagraphElement) {
                var wordParagraph = new WordParagraph(source.Document, wordParagraphElement);
                RtfParagraph paragraph = destination.AddParagraph();
                CopyTabStops(wordParagraph, paragraph);
                CopyParagraphFormatting(wordParagraph, paragraph, document);
                AppendFormattedRuns(wordParagraph, paragraph, document, revisionAuthorIndexes);
            } else if (child is Table nestedTableElement) {
                var nestedWordTable = new WordTable(source.Document, nestedTableElement);
                RtfTable nested = destination.AddTable(0, GetColumnCount(nestedWordTable));
                CopyTable(nestedWordTable, nested, document, revisionAuthorIndexes, tableDepth + 1);
            }
        }

        if (destination.Blocks.Count == 0) {
            destination.AddParagraph();
        }
    }

    private static int GetColumnCount(WordTable table) {
        int columnCount = table.Rows.Select(row => row.GetCells(readOnly: true).Count).DefaultIfEmpty(0).Max();
        return Math.Max(1, columnCount);
    }

    private static void AppendTable(WordDocument document, RtfTable source, RtfDocument rtfDocument) {
        WordTable table = document.AddTable(GetRowCount(source), GetColumnCount(source));
        ApplyTable(source, table, rtfDocument);
    }

    private static void AppendTable(WordSection section, RtfTable source, RtfDocument rtfDocument) {
        WordTable table = section.AddTable(GetRowCount(source), GetColumnCount(source));
        ApplyTable(source, table, rtfDocument);
    }

    private static void ApplyTable(RtfTable source, WordTable destination, RtfDocument rtfDocument) {
        ApplyUniformTableCellSpacing(source, destination);
        ApplyUniformTableLeftIndent(source, destination);
        ApplyUniformTableAlignment(source, destination);
        ApplyUniformTablePreferredWidth(source, destination);
        for (int rowIndex = 0; rowIndex < source.Rows.Count; rowIndex++) {
            RtfTableRow sourceRow = source.Rows[rowIndex];
            WordTableRow destinationRow = destination.Rows[rowIndex];
            destinationRow.RepeatHeaderRowAtTheTopOfEachPage = sourceRow.RepeatHeader;
            destinationRow.AllowRowToBreakAcrossPages = !sourceRow.KeepTogether;
            destinationRow.Height = sourceRow.HeightTwips;
            int previousBoundary = 0;
            for (int cellIndex = 0; cellIndex < sourceRow.Cells.Count && cellIndex < destinationRow.Cells.Count; cellIndex++) {
                RtfTableCell sourceCell = sourceRow.Cells[cellIndex];
                WordTableCell destinationCell = destinationRow.Cells[cellIndex];
                ApplyCellWidth(sourceCell, destinationCell, ref previousBoundary);
                ApplyCellProperties(sourceCell, destinationCell, rtfDocument);
                ApplyCellParagraphs(sourceCell, destinationCell, rtfDocument);
            }
        }
    }

    private static int? GetWordTableCellSpacing(WordTable table) {
        string? width = table._tableProperties?.TableCellSpacing?.Width?.Value;
        if (int.TryParse(width, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)) {
            return value;
        }

        return null;
    }

    private static int? GetWordTableLeftIndent(WordTable table) {
        TableIndentation? indentation = table._tableProperties?.TableIndentation;
        if (indentation?.Type?.Value != TableWidthUnitValues.Dxa) {
            return null;
        }

        return indentation.Width?.Value;
    }

    private static void ApplyUniformTableCellSpacing(RtfTable source, WordTable destination) {
        if (source.Rows.Count == 0 || source.Rows.Any(row => !row.CellGapTwips.HasValue)) {
            return;
        }

        int firstGap = source.Rows[0].CellGapTwips!.Value;
        if (source.Rows.Any(row => row.CellGapTwips!.Value != firstGap)) {
            return;
        }

        destination.CheckTableProperties();
        destination._tableProperties!.TableCellSpacing = new TableCellSpacing {
            Width = firstGap.ToString(CultureInfo.InvariantCulture),
            Type = TableWidthUnitValues.Dxa
        };
    }

    private static void ApplyUniformTableLeftIndent(RtfTable source, WordTable destination) {
        if (source.Rows.Count == 0 || source.Rows.Any(row => !row.LeftIndentTwips.HasValue)) {
            return;
        }

        int firstIndent = source.Rows[0].LeftIndentTwips!.Value;
        if (source.Rows.Any(row => row.LeftIndentTwips!.Value != firstIndent)) {
            return;
        }

        destination.CheckTableProperties();
        destination._tableProperties!.TableIndentation = new TableIndentation {
            Width = firstIndent,
            Type = TableWidthUnitValues.Dxa
        };
    }

    private static void ApplyUniformTableAlignment(RtfTable source, WordTable destination) {
        if (source.Rows.Count == 0 || source.Rows.Any(row => !row.Alignment.HasValue)) {
            return;
        }

        RtfTableAlignment firstAlignment = source.Rows[0].Alignment!.Value;
        if (source.Rows.Any(row => row.Alignment!.Value != firstAlignment)) {
            return;
        }

        destination.Alignment = ToWordTableAlignment(firstAlignment);
    }

    private static void ApplyUniformTablePreferredWidth(RtfTable source, WordTable destination) {
        if (source.Rows.Count == 0 || source.Rows.Any(row => !row.PreferredWidth.HasValue && !row.PreferredWidthUnit.HasValue)) {
            return;
        }

        int? firstWidth = source.Rows[0].PreferredWidth;
        RtfTableWidthUnit? firstUnit = source.Rows[0].PreferredWidthUnit;
        if (source.Rows.Any(row => row.PreferredWidth != firstWidth || row.PreferredWidthUnit != firstUnit)) {
            return;
        }

        TableWidthUnitValues? wordUnit = ToWordTableWidthUnit(firstUnit);
        if (wordUnit.HasValue) {
            destination.WidthType = wordUnit.Value;
        }

        if (firstWidth.HasValue) {
            destination.Width = firstWidth.Value;
        }
    }

    private static void ApplyCellWidth(RtfTableCell source, WordTableCell destination, ref int previousBoundary) {
        if (!source.RightBoundaryTwips.HasValue) {
            return;
        }

        int width = source.RightBoundaryTwips.Value - previousBoundary;
        previousBoundary = source.RightBoundaryTwips.Value;
        if (width <= 0) {
            return;
        }

        destination.WidthType = TableWidthUnitValues.Dxa;
        destination.Width = width;
    }

    private static void ApplyCellProperties(RtfTableCell source, WordTableCell destination, RtfDocument document) {
        destination.HorizontalMerge = ToWordMergedCellValue(source.HorizontalMerge);
        destination.VerticalMerge = ToWordMergedCellValue(source.VerticalMerge);
        destination.VerticalAlignment = ToWordCellVerticalAlignment(source.VerticalAlignment);
        destination.TextDirection = ToWordCellTextDirection(source.TextFlow);
        ApplyCellPreferredWidth(source, destination);
        destination.WrapText = !source.NoWrap;
        destination.FitText = source.FitText;
        if (source.BackgroundColorIndex.HasValue) {
            string? color = GetColorHex(document, source.BackgroundColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                destination.ShadingFillColorHex = color!;
            }
        }

        Shading? shading = null;
        if (source.ShadingForegroundColorIndex.HasValue) {
            string? color = GetColorHex(document, source.ShadingForegroundColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                shading = GetOrCreateCellShading(destination);
                shading.Color = color!;
            }
        }

        ShadingPatternValues? pattern = ToWordShadingPercent(source.ShadingPatternPercent) ?? ToWordShadingPattern(source.ShadingPattern);
        if (pattern.HasValue) {
            shading = GetOrCreateCellShading(destination);
            shading.Val = pattern.Value;
        }

        ApplyCellPadding(source, destination);
        ApplyCellBorders(source, destination, document);
    }

    private static void ApplyCellPreferredWidth(RtfTableCell source, WordTableCell destination) {
        TableWidthUnitValues? wordUnit = ToWordTableWidthUnit(source.PreferredWidthUnit);
        if (wordUnit.HasValue) {
            destination.WidthType = wordUnit.Value;
        }

        if (source.PreferredWidth.HasValue) {
            destination.Width = source.PreferredWidth.Value;
        }
    }

    private static Shading GetOrCreateCellShading(WordTableCell cell) {
        cell.AddTableCellProperties();
        cell._tableCellProperties!.Shading ??= new Shading();
        return cell._tableCellProperties.Shading;
    }

    private static void ApplyCellPadding(RtfTableCell source, WordTableCell destination) {
        if (TryToInt16(source.PaddingTopTwips, out short top)) {
            destination.MarginTopWidth = top;
        }

        if (TryToInt16(source.PaddingLeftTwips, out short left)) {
            destination.MarginLeftWidth = left;
        }

        if (TryToInt16(source.PaddingBottomTwips, out short bottom)) {
            destination.MarginBottomWidth = bottom;
        }

        if (TryToInt16(source.PaddingRightTwips, out short right)) {
            destination.MarginRightWidth = right;
        }
    }

    private static void ApplyCellBorders(RtfTableCell source, WordTableCell destination, RtfDocument document) {
        ApplyCellBorder(source.TopBorder, style => destination.Borders.TopStyle = style, width => destination.Borders.TopSize = width, color => destination.Borders.TopColorHex = color, document);
        ApplyCellBorder(source.LeftBorder, style => destination.Borders.LeftStyle = style, width => destination.Borders.LeftSize = width, color => destination.Borders.LeftColorHex = color, document);
        ApplyCellBorder(source.BottomBorder, style => destination.Borders.BottomStyle = style, width => destination.Borders.BottomSize = width, color => destination.Borders.BottomColorHex = color, document);
        ApplyCellBorder(source.RightBorder, style => destination.Borders.RightStyle = style, width => destination.Borders.RightSize = width, color => destination.Borders.RightColorHex = color, document);
        ApplyCellBorder(source.TopLeftToBottomRightBorder, style => destination.Borders.TopLeftToBottomRightStyle = style, width => destination.Borders.TopLeftToBottomRightSize = width, color => destination.Borders.TopLeftToBottomRightColorHex = color, document);
        ApplyCellBorder(source.TopRightToBottomLeftBorder, style => destination.Borders.TopRightToBottomLeftStyle = style, width => destination.Borders.TopRightToBottomLeftSize = width, color => destination.Borders.TopRightToBottomLeftColorHex = color, document);
    }

    private static void ApplyCellBorder(RtfTableCellBorder source, Action<BorderValues?> setStyle, Action<UInt32Value?> setWidth, Action<string?> setColor, RtfDocument document) {
        if (!source.HasAnyValue) {
            return;
        }

        setStyle(ToWordBorderStyle(source.Style));
        if (source.Width.HasValue && source.Width.Value >= 0) {
            setWidth((UInt32Value)(uint)source.Width.Value);
        }

        if (source.ColorIndex.HasValue) {
            string? color = GetColorHex(document, source.ColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                setColor(color);
            }
        }
    }

    private static void ApplyCellParagraphs(RtfTableCell source, WordTableCell destination, RtfDocument rtfDocument) {
        if (source.Blocks.Count == 0) {
            destination.AddParagraph(removeExistingParagraphs: true);
            return;
        }

        bool first = true;
        foreach (IRtfBlock block in source.Blocks) {
            if (block is RtfParagraph sourceParagraph) {
                WordParagraph paragraph = destination.AddParagraph(removeExistingParagraphs: first);
                ApplyTabStops(paragraph, sourceParagraph);
                ApplyParagraphFormatting(paragraph, sourceParagraph, rtfDocument);
                AppendRuns(paragraph, sourceParagraph, rtfDocument);
                first = false;
            } else if (block is RtfTable nestedTable) {
                WordTable wordNestedTable = destination.AddTable(GetRowCount(nestedTable), GetColumnCount(nestedTable), WordTableStyle.TableGrid, removePrecedingParagraph: first);
                ApplyTable(nestedTable, wordNestedTable, rtfDocument);
                first = false;
            }
        }
    }

    private static int GetRowCount(RtfTable table) {
        return Math.Max(1, table.Rows.Count);
    }

    private static int GetColumnCount(RtfTable table) {
        int columnCount = table.Rows.Select(row => row.Cells.Count).DefaultIfEmpty(0).Max();
        return Math.Max(1, columnCount);
    }

    private static RtfTableCellMerge ToRtfCellMerge(MergedCellValues? value) {
        if (value == MergedCellValues.Restart) return RtfTableCellMerge.First;
        if (value == MergedCellValues.Continue) return RtfTableCellMerge.Continue;
        return RtfTableCellMerge.None;
    }

    private static MergedCellValues? ToWordMergedCellValue(RtfTableCellMerge merge) {
        if (merge == RtfTableCellMerge.First) return MergedCellValues.Restart;
        if (merge == RtfTableCellMerge.Continue) return MergedCellValues.Continue;
        return null;
    }

    private static RtfTableCellVerticalAlignment? ToRtfCellVerticalAlignment(TableVerticalAlignmentValues? value) {
        if (value == TableVerticalAlignmentValues.Center) return RtfTableCellVerticalAlignment.Center;
        if (value == TableVerticalAlignmentValues.Bottom) return RtfTableCellVerticalAlignment.Bottom;
        if (value == TableVerticalAlignmentValues.Top) return RtfTableCellVerticalAlignment.Top;
        return null;
    }

    private static TableVerticalAlignmentValues? ToWordCellVerticalAlignment(RtfTableCellVerticalAlignment? value) {
        if (value == RtfTableCellVerticalAlignment.Center) return TableVerticalAlignmentValues.Center;
        if (value == RtfTableCellVerticalAlignment.Bottom) return TableVerticalAlignmentValues.Bottom;
        if (value == RtfTableCellVerticalAlignment.Top) return TableVerticalAlignmentValues.Top;
        return null;
    }

    private static RtfTableCellTextFlow? ToRtfCellTextFlow(TextDirectionValues? value) {
        if (value == TextDirectionValues.LefToRightTopToBottom) return RtfTableCellTextFlow.LeftToRightTopToBottom;
        if (value == TextDirectionValues.TopToBottomRightToLeft) return RtfTableCellTextFlow.TopToBottomRightToLeft;
        if (value == TextDirectionValues.BottomToTopLeftToRight) return RtfTableCellTextFlow.BottomToTopLeftToRight;
        if (value == TextDirectionValues.LefttoRightTopToBottomRotated) return RtfTableCellTextFlow.LeftToRightTopToBottomVertical;
        if (value == TextDirectionValues.TopToBottomRightToLeftRotated) return RtfTableCellTextFlow.TopToBottomRightToLeftVertical;
        return null;
    }

    private static TextDirectionValues? ToWordCellTextDirection(RtfTableCellTextFlow? value) {
        if (value == RtfTableCellTextFlow.LeftToRightTopToBottom) return TextDirectionValues.LefToRightTopToBottom;
        if (value == RtfTableCellTextFlow.TopToBottomRightToLeft) return TextDirectionValues.TopToBottomRightToLeft;
        if (value == RtfTableCellTextFlow.BottomToTopLeftToRight) return TextDirectionValues.BottomToTopLeftToRight;
        if (value == RtfTableCellTextFlow.LeftToRightTopToBottomVertical) return TextDirectionValues.LefttoRightTopToBottomRotated;
        if (value == RtfTableCellTextFlow.TopToBottomRightToLeftVertical) return TextDirectionValues.TopToBottomRightToLeftRotated;
        return null;
    }

    private static RtfTableAlignment? ToRtfTableAlignment(TableRowAlignmentValues? value) {
        if (value == TableRowAlignmentValues.Center) return RtfTableAlignment.Center;
        if (value == TableRowAlignmentValues.Right) return RtfTableAlignment.Right;
        if (value == TableRowAlignmentValues.Left) return RtfTableAlignment.Left;
        return null;
    }

    private static RtfTableWidthUnit? ToRtfTableWidthUnit(TableWidthUnitValues? value) {
        if (value == TableWidthUnitValues.Auto) return RtfTableWidthUnit.Auto;
        if (value == TableWidthUnitValues.Dxa) return RtfTableWidthUnit.Twips;
        if (value == TableWidthUnitValues.Pct) return RtfTableWidthUnit.Percent;
        return null;
    }

    private static TableRowAlignmentValues? ToWordTableAlignment(RtfTableAlignment? value) {
        if (value == RtfTableAlignment.Center) return TableRowAlignmentValues.Center;
        if (value == RtfTableAlignment.Right) return TableRowAlignmentValues.Right;
        if (value == RtfTableAlignment.Left) return TableRowAlignmentValues.Left;
        return null;
    }

    private static TableWidthUnitValues? ToWordTableWidthUnit(RtfTableWidthUnit? value) {
        if (value == RtfTableWidthUnit.Auto) return TableWidthUnitValues.Auto;
        if (value == RtfTableWidthUnit.Twips) return TableWidthUnitValues.Dxa;
        if (value == RtfTableWidthUnit.Percent) return TableWidthUnitValues.Pct;
        return null;
    }

    private static RtfTableCellBorderStyle ToRtfBorderStyle(BorderValues? value) {
        if (value == BorderValues.Double) return RtfTableCellBorderStyle.Double;
        if (value == BorderValues.Dotted) return RtfTableCellBorderStyle.Dotted;
        if (value == BorderValues.Dashed) return RtfTableCellBorderStyle.Dashed;
        if (value == BorderValues.Nil || value == BorderValues.None) return RtfTableCellBorderStyle.None;
        if (value == BorderValues.Single) return RtfTableCellBorderStyle.Single;
        return RtfTableCellBorderStyle.None;
    }

    private static BorderValues? ToWordBorderStyle(RtfTableCellBorderStyle value) {
        switch (value) {
            case RtfTableCellBorderStyle.Double:
                return BorderValues.Double;
            case RtfTableCellBorderStyle.Dotted:
                return BorderValues.Dotted;
            case RtfTableCellBorderStyle.Dashed:
                return BorderValues.Dashed;
            case RtfTableCellBorderStyle.Single:
                return BorderValues.Single;
            default:
                return BorderValues.Nil;
        }
    }

    private static bool TryToInt16(int? value, out short result) {
        result = 0;
        if (!value.HasValue || value.Value < short.MinValue || value.Value > short.MaxValue) {
            return false;
        }

        result = (short)value.Value;
        return true;
    }

    private static int GetOrAddColor(RtfDocument document, byte red, byte green, byte blue) {
        for (int index = 0; index < document.Colors.Count; index++) {
            RtfColor color = document.Colors[index];
            if (color.Red == red && color.Green == green && color.Blue == blue) {
                return index + 1;
            }
        }

        return document.AddColor(red, green, blue);
    }

    private static string? GetColorHex(RtfDocument document, int colorIndex) {
        if (colorIndex <= 0 || colorIndex > document.Colors.Count) {
            return null;
        }

        RtfColor color = document.Colors[colorIndex - 1];
        return $"{color.Red:X2}{color.Green:X2}{color.Blue:X2}";
    }

    private static bool TryParseHexColor(string value, out byte red, out byte green, out byte blue) {
        string hex = value.Trim().TrimStart('#');
        red = 0;
        green = 0;
        blue = 0;
        if (hex.Length != 6) {
            return false;
        }

        return byte.TryParse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out red) &&
            byte.TryParse(hex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out green) &&
            byte.TryParse(hex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out blue);
    }
}
