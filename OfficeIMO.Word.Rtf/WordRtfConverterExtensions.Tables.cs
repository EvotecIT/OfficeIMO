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
                CopyParagraphFormatting(wordParagraph, paragraph, document);
                AppendFormattedRuns(wordParagraph, paragraph, document, revisionAuthorIndexes);
            } else if (element is WordTable table) {
                CopyTable(table, destination.AddTable(0, GetColumnCount(table)), document, revisionAuthorIndexes, 1);
            }
        }
    }

    private static void CopyTable(WordTable source, RtfTable destination, RtfDocument document, Dictionary<string, int> revisionAuthorIndexes, int tableDepth) {
        RtfTableTraversalGuard.EnsureDepth(tableDepth);
        int? cellGap = GetWordTableCellGap(source);
        int? tableLeftIndent = GetWordTableLeftIndent(source);
        if (cellGap.HasValue) {
            tableLeftIndent = (tableLeftIndent ?? 0) - GetWordTableGapIndentAdjustment(source, cellGap.Value);
        }
        int? tableLeftPadding = cellGap.HasValue ? null : GetWordTableLeftPadding(source);
        int? tableRightPadding = cellGap.HasValue ? null : GetWordTableRightPadding(source);
        RtfTableAlignment? tableAlignment = ToRtfTableAlignment(source.Alignment.ToOpenXml());
        RtfTableWidthUnit? tableWidthUnit = ToRtfTableWidthUnit(source._tableProperties?.TableWidth);
        int? tableWidth = ToRtfPreferredWidthValue(source._tableProperties?.TableWidth);
        foreach (WordTableRow wordRow in source.Rows) {
            RtfTableRow row = destination.AddRow();
            row.RepeatHeader = wordRow.RepeatHeaderRowAtTheTopOfEachPage;
            row.KeepTogether = !wordRow.AllowRowToBreakAcrossPages;
            row.HeightTwips = checked((int?)(wordRow.Height?.ToTwips()));
            row.CellGapTwips = cellGap;
            row.LeftIndentTwips = tableLeftIndent;
            row.PaddingLeftTwips = tableLeftPadding;
            row.PaddingRightTwips = tableRightPadding;
            row.Alignment = tableAlignment;
            row.PreferredWidthUnit = tableWidthUnit;
            row.PreferredWidth = tableWidth;
            int rightBoundary = 0;
            foreach (WordTableCell wordCell in wordRow.GetCells(readOnly: true)) {
                int width = wordCell.Width?.LengthValue is { } physicalWidth
                    ? Math.Max(1, checked((int)physicalWidth.ToTwips()))
                    : 2400;
                rightBoundary += width;
                RtfTableCell cell = row.AddCell(rightBoundary);
                CopyCellProperties(wordCell, cell, document);
                CopyCellParagraphs(wordCell, cell, document, revisionAuthorIndexes, tableDepth);
            }
        }
    }

    private static void CopyCellProperties(WordTableCell source, RtfTableCell destination, RtfDocument document) {
        destination.HorizontalMerge = ToRtfCellMerge(source.HorizontalMerge.ToOpenXml());
        destination.VerticalMerge = ToRtfCellMerge(source.VerticalMerge.ToOpenXml());
        destination.VerticalAlignment = ToRtfCellVerticalAlignment(source.VerticalAlignment.ToOpenXml());
        destination.TextFlow = ToRtfCellTextFlow(source.TextDirection.ToOpenXml());
        destination.PreferredWidthUnit = ToRtfTableWidthUnit(source._tableCellProperties?.TableCellWidth);
        destination.PreferredWidth = ToRtfPreferredWidthValue(source._tableCellProperties?.TableCellWidth);
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

        CopyCellShadingPattern(source.ShadingPattern.ToOpenXml(), destination);
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
        destination.PaddingTopTwips = checked((short?)(source.MarginTop?.ToTwips()));
        destination.PaddingLeftTwips = checked((short?)(source.MarginLeft?.ToTwips()));
        destination.PaddingBottomTwips = checked((short?)(source.MarginBottom?.ToTwips()));
        destination.PaddingRightTwips = checked((short?)(source.MarginRight?.ToTwips()));
    }

    private static void CopyCellBorders(WordTableCell source, RtfTableCell destination, RtfDocument document) {
        CopyCellBorder(source.Borders.TopStyle.ToOpenXml(), WordBorderLengths.ToSize(source.Borders.TopSize), source.Borders.TopColorHex, destination.TopBorder, document);
        CopyCellBorder(source.Borders.LeftStyle.ToOpenXml(), WordBorderLengths.ToSize(source.Borders.LeftSize), source.Borders.LeftColorHex, destination.LeftBorder, document);
        CopyCellBorder(source.Borders.BottomStyle.ToOpenXml(), WordBorderLengths.ToSize(source.Borders.BottomSize), source.Borders.BottomColorHex, destination.BottomBorder, document);
        CopyCellBorder(source.Borders.RightStyle.ToOpenXml(), WordBorderLengths.ToSize(source.Borders.RightSize), source.Borders.RightColorHex, destination.RightBorder, document);
        CopyCellBorder(source.Borders.TopLeftToBottomRightStyle.ToOpenXml(), WordBorderLengths.ToSize(source.Borders.TopLeftToBottomRightSize), source.Borders.TopLeftToBottomRightColorHex, destination.TopLeftToBottomRightBorder, document);
        CopyCellBorder(source.Borders.TopRightToBottomLeftStyle.ToOpenXml(), WordBorderLengths.ToSize(source.Borders.TopRightToBottomLeftSize), source.Borders.TopRightToBottomLeftColorHex, destination.TopRightToBottomLeftBorder, document);
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

    private static void AppendTable(WordDocument document, RtfTable source, RtfDocument rtfDocument, RtfWordNumberingPlan numbering) {
        WordTable table = document.AddTable(GetRowCount(source), GetColumnCount(source));
        ApplyTable(source, table, rtfDocument, numbering);
    }

    private static void AppendTable(WordSection section, RtfTable source, RtfDocument rtfDocument, RtfWordNumberingPlan numbering) {
        WordTable table = section.AddTable(GetRowCount(source), GetColumnCount(source));
        ApplyTable(source, table, rtfDocument, numbering);
    }

    private static void ApplyTable(RtfTable source, WordTable destination, RtfDocument rtfDocument, RtfWordNumberingPlan numbering) {
        destination.Style = WordTableStyle.TableNormal;
        destination.StyleDetails!.SetCustomBorders(topStyle: WordBorderStyle.Nil, bottomStyle: WordBorderStyle.Nil,
            leftStyle: WordBorderStyle.Nil, rightStyle: WordBorderStyle.Nil,
            insideHStyle: WordBorderStyle.Nil, insideVStyle: WordBorderStyle.Nil);
        int? cellGap = GetUniformTableCellGap(source);
        ApplyTableCellGap(destination, cellGap);
        ApplyUniformTableLeftIndent(source, destination, cellGap);
        ApplyUniformTableAlignment(source, destination);
        ApplyUniformTablePreferredWidth(source, destination);
        for (int rowIndex = 0; rowIndex < source.Rows.Count; rowIndex++) {
            RtfTableRow sourceRow = source.Rows[rowIndex];
            WordTableRow destinationRow = destination.Rows[rowIndex];
            destinationRow.RepeatHeaderRowAtTheTopOfEachPage = sourceRow.RepeatHeader;
            destinationRow.AllowRowToBreakAcrossPages = !sourceRow.KeepTogether;
            destinationRow.Height= ((sourceRow.HeightTwips) is { } wordLength11622 ? global::OfficeIMO.OfficeLength.Points((wordLength11622) / 20D) : (global::OfficeIMO.OfficeLength?)null);
            int previousBoundary = 0;
            for (int cellIndex = 0; cellIndex < sourceRow.Cells.Count && cellIndex < destinationRow.Cells.Count; cellIndex++) {
                RtfTableCell sourceCell = sourceRow.Cells[cellIndex];
                WordTableCell destinationCell = destinationRow.Cells[cellIndex];
                ApplyCellWidth(sourceCell, destinationCell, ref previousBoundary);
                ApplyCellProperties(sourceCell, sourceRow, destinationCell, rtfDocument);
                ApplyCellParagraphs(sourceCell, destinationCell, rtfDocument, numbering);
            }
        }
    }

    private static int? GetWordTableLeftIndent(WordTable table) {
        TableIndentation? indentation = table._tableProperties?.TableIndentation;
        if (indentation?.Type?.Value != TableWidthUnitValues.Dxa) {
            return null;
        }

        return indentation.Width?.Value;
    }

    /// <summary>Recovers an RTF half-gap only from symmetric direct dxa margins and a reversible indent.</summary>
    private static int? GetWordTableCellGap(WordTable table) {
        int? left = GetWordTableLeftPadding(table);
        int? right = GetWordTableRightPadding(table);
        if (!left.HasValue || left.Value < 0 || left != right) {
            return null;
        }

        TableIndentation? indentation = table._tableProperties?.TableIndentation;
        if (indentation != null &&
            (indentation.Type?.Value != TableWidthUnitValues.Dxa || indentation.Width == null)) {
            return null;
        }
        long sourceIndent = (long)(indentation?.Width?.Value ?? 0) - GetWordTableGapIndentAdjustment(table, left.Value);
        return sourceIndent >= int.MinValue && sourceIndent <= int.MaxValue ? left : null;
    }

    /// <summary>Legacy Word placement offsets the inline grid by its leading cell margin; modern placement does not.</summary>
    private static int GetWordTableGapIndentAdjustment(WordTable table, int gap) =>
        table.Document.SourceFormat == WordFileFormat.Doc ||
        (int)table.Document.CompatibilitySettings.CompatibilityMode < (int)WordCompatibilityMode.Word2013
            ? gap : 0;

    private static int? GetWordTableLeftPadding(WordTable table) {
        TableCellLeftMargin? margin = table._tableProperties?.TableCellMarginDefault?.TableCellLeftMargin;
        return margin?.Type?.Value == TableWidthValues.Dxa ? margin.Width?.Value : null;
    }

    private static int? GetWordTableRightPadding(WordTable table) {
        TableCellRightMargin? margin = table._tableProperties?.TableCellMarginDefault?.TableCellRightMargin;
        return margin?.Type?.Value == TableWidthValues.Dxa ? margin.Width?.Value : null;
    }

    /// <summary>Only an explicit uniform nonnegative gap can become a shared Word table margin.</summary>
    private static int? GetUniformTableCellGap(RtfTable source) {
        if (source.Rows.Count == 0 || source.Rows.Any(row => !row.CellGapTwips.HasValue)) {
            return null;
        }
        int gap = source.Rows[0].CellGapTwips!.Value;
        if (gap < 0 || !TryToInt16(gap, out _) || source.Rows.Any(row => row.CellGapTwips != gap)) {
            return null;
        }
        return gap;
    }

    /// <summary>RTF row gaps are cell margins, not Word spacing between separated cells.</summary>
    private static void ApplyTableCellGap(WordTable destination, int? gap) {
        destination.CheckTableProperties();
        destination._tableProperties!.TableCellSpacing?.Remove();
        TableCellMarginDefault margins = destination._tableProperties.TableCellMarginDefault ??= new TableCellMarginDefault();
        short width = (short)(gap ?? 0);
        margins.TableCellLeftMargin = new TableCellLeftMargin { Width = width, Type = TableWidthValues.Dxa };
        margins.TableCellRightMargin = new TableCellRightMargin { Width = width, Type = TableWidthValues.Dxa };
    }

    private static void ApplyUniformTableLeftIndent(RtfTable source, WordTable destination, int? gap) {
        if (source.Rows.Count == 0) {
            return;
        }
        int? sourceIndent = source.Rows[0].LeftIndentTwips;
        if (source.Rows.Any(row => row.LeftIndentTwips != sourceIndent) ||
            (!sourceIndent.HasValue && !gap.HasValue)) {
            return;
        }
        long indent = (long)(sourceIndent ?? 0) + GetWordTableGapIndentAdjustment(destination, gap ?? 0);
        if (indent < int.MinValue || indent > int.MaxValue) {
            return;
        }

        destination.CheckTableProperties();
        if (indent == 0) {
            destination._tableProperties!.TableIndentation?.Remove();
            return;
        }
        destination._tableProperties!.TableIndentation = new TableIndentation {
            Width = (int)indent,
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

        destination.Alignment = ToWordTableAlignment(firstAlignment).ToOfficeEnum();
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

        ApplyNativeTablePreference(destination, firstUnit, firstWidth);
    }

    // Preserve partial RTF declarations rather than inventing a missing measurement or unit.
    private static void ApplyNativeTablePreference(WordTable destination, RtfTableWidthUnit? unit, int? value) {
        destination.CheckTableProperties();
        var native = destination._tableProperties!.TableWidth?.CloneNode(true) as TableWidth ?? new TableWidth();
        var type = ToWordTableWidthUnit(unit);
        if (type.HasValue) native.Type = type.Value;
        if (value.HasValue) native.Width = value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        destination._tableProperties!.TableWidth = native;
        destination.RefreshGrid();
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

        destination.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus((long)width * 635));
    }

    private static void ApplyCellProperties(RtfTableCell source, RtfTableRow row, WordTableCell destination, RtfDocument document) {
        destination.HorizontalMerge = ToWordMergedCellValue(source.HorizontalMerge).ToOfficeEnum();
        destination.VerticalMerge = ToWordMergedCellValue(source.VerticalMerge).ToOfficeEnum();
        destination.VerticalAlignment = ToWordCellVerticalAlignment(source.VerticalAlignment).ToOfficeEnum();
        destination.TextDirection = ToWordCellTextDirection(source.TextFlow).ToOfficeEnum();
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

        ApplyCellPadding(source, row, destination);
        ApplyCellBorders(source, destination, document);
    }

    private static void ApplyCellPreferredWidth(RtfTableCell source, WordTableCell destination) {
        if (!source.PreferredWidthUnit.HasValue && !source.PreferredWidth.HasValue) return;
        destination.AddTableCellProperties();
        var native = destination._tableCellProperties!.TableCellWidth?.CloneNode(true) as TableCellWidth ?? new TableCellWidth();
        var type = ToWordTableWidthUnit(source.PreferredWidthUnit);
        if (type.HasValue) native.Type = type.Value;
        if (source.PreferredWidth.HasValue) native.Width = source.PreferredWidth.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        destination._tableCellProperties!.TableCellWidth = native;
    }

    private static Shading GetOrCreateCellShading(WordTableCell cell) {
        cell.AddTableCellProperties();
        cell._tableCellProperties!.Shading ??= new Shading();
        return cell._tableCellProperties.Shading;
    }

    private static void ApplyCellPadding(RtfTableCell source, RtfTableRow row, WordTableCell destination) {
        if (TryToInt16(source.PaddingTopTwips ?? row.PaddingTopTwips ?? 0, out short top)) {
            destination.MarginTop = global::OfficeIMO.OfficeLength.Points((top) / 20D);
        }
        if (TryToInt16(source.PaddingBottomTwips ?? row.PaddingBottomTwips ?? 0, out short bottom)) {
            destination.MarginBottom = global::OfficeIMO.OfficeLength.Points((bottom) / 20D);
        }

        // Omitted horizontal padding inherits the table gap; explicit cell values, including zero, override row values.
        destination.MarginLeft= ((TryToInt16(source.PaddingLeftTwips ?? row.PaddingLeftTwips, out short left)
            ? left : (short?)null) is { } wordLength21241 ? global::OfficeIMO.OfficeLength.Points((wordLength21241) / 20D) : (global::OfficeIMO.OfficeLength?)null);
        destination.MarginRight= ((TryToInt16(source.PaddingRightTwips ?? row.PaddingRightTwips, out short right)
            ? right : (short?)null) is { } wordLength21391 ? global::OfficeIMO.OfficeLength.Points((wordLength21391) / 20D) : (global::OfficeIMO.OfficeLength?)null);
    }

    private static void ApplyCellBorders(RtfTableCell source, WordTableCell destination, RtfDocument document) {
        ApplyCellBorder(source.TopBorder, style => destination.Borders.TopStyle = style.ToOfficeEnum(), width => destination.Borders.TopSize = WordBorderLengths.FromSize(width), color => destination.Borders.TopColorHex = color, document);
        ApplyCellBorder(source.LeftBorder, style => destination.Borders.LeftStyle = style.ToOfficeEnum(), width => destination.Borders.LeftSize = WordBorderLengths.FromSize(width), color => destination.Borders.LeftColorHex = color, document);
        ApplyCellBorder(source.BottomBorder, style => destination.Borders.BottomStyle = style.ToOfficeEnum(), width => destination.Borders.BottomSize = WordBorderLengths.FromSize(width), color => destination.Borders.BottomColorHex = color, document);
        ApplyCellBorder(source.RightBorder, style => destination.Borders.RightStyle = style.ToOfficeEnum(), width => destination.Borders.RightSize = WordBorderLengths.FromSize(width), color => destination.Borders.RightColorHex = color, document);
        ApplyCellBorder(source.TopLeftToBottomRightBorder, style => destination.Borders.TopLeftToBottomRightStyle = style.ToOfficeEnum(), width => destination.Borders.TopLeftToBottomRightSize = WordBorderLengths.FromSize(width), color => destination.Borders.TopLeftToBottomRightColorHex = color, document);
        ApplyCellBorder(source.TopRightToBottomLeftBorder, style => destination.Borders.TopRightToBottomLeftStyle = style.ToOfficeEnum(), width => destination.Borders.TopRightToBottomLeftSize = WordBorderLengths.FromSize(width), color => destination.Borders.TopRightToBottomLeftColorHex = color, document);
    }

    private static void ApplyCellBorder(RtfTableCellBorder source, Action<BorderValues?> setStyle, Action<uint?> setWidth, Action<string?> setColor, RtfDocument document) {
        if (!source.HasAnyValue) {
            return;
        }

        setStyle(ToWordBorderStyle(source.Style));
        if (source.Width.HasValue && source.Width.Value >= 0) {
            setWidth((uint)source.Width.Value);
        }

        if (source.ColorIndex.HasValue) {
            string? color = GetColorHex(document, source.ColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                setColor(color);
            }
        }
    }

    private static void ApplyCellParagraphs(RtfTableCell source, WordTableCell destination, RtfDocument rtfDocument, RtfWordNumberingPlan numbering) {
        if (source.Blocks.Count == 0) {
            destination.AddParagraph(removeExistingParagraphs: true);
            return;
        }

        bool first = true;
        foreach (IRtfBlock block in source.Blocks) {
            if (block is RtfParagraph sourceParagraph) {
                WordParagraph paragraph = destination.AddParagraph(removeExistingParagraphs: first);
                ApplyParagraphFormatting(paragraph, sourceParagraph, rtfDocument, numbering);
                AppendRuns(paragraph, sourceParagraph, rtfDocument, numbering);
                first = false;
            } else if (block is RtfTable nestedTable) {
                WordTable wordNestedTable = destination.AddTable(GetRowCount(nestedTable), GetColumnCount(nestedTable), WordTableStyle.TableGrid, removePrecedingParagraph: first);
                ApplyTable(nestedTable, wordNestedTable, rtfDocument, numbering);
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

    private static RtfTableWidthUnit? ToRtfTableWidthUnit(TableWidthType? width) {
        if (width == null) return null;
        var type = WordTableWidthNative.NativeType(width);
        if (!type.HasValue) throw new NotSupportedException("Unsupported Word preferred-width type: " + width.Type?.InnerText);
        // Lexical percentages remain relative in RTF even if the native unit was omitted.
        if (width.Width?.Value?.Trim().EndsWith("%", StringComparison.Ordinal) == true) return RtfTableWidthUnit.Percent;
        if (width.Type == null) return null;
        if (type == TableWidthUnitValues.Auto) return RtfTableWidthUnit.Auto;
        if (type == TableWidthUnitValues.Dxa) return RtfTableWidthUnit.Twips;
        if (type == TableWidthUnitValues.Pct) return RtfTableWidthUnit.Percent;
        return null;
    }

    private static int? ToRtfPreferredWidthValue(TableWidthType? width) {
        if (width?.Width == null) return null;
        return WordTableWidthNative.NumericValue(width) ?? WordTableWidthNative.Read(width)?.ToNativeValue();
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
            if (!color.IsAutomatic && color.Red == red && color.Green == green && color.Blue == blue) {
                return index + 1;
            }
        }

        return document.AddColor(red, green, blue);
    }

    private static string? GetColorHex(RtfDocument document, int colorIndex) {
        RtfColor? color = document.GetColor(colorIndex);
        return color == null ? null : $"{color.Red:X2}{color.Green:X2}{color.Blue:X2}";
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
