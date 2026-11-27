using System.Globalization;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyWordElements(IEnumerable<WordElement> elements, RtfDocument destination, Dictionary<string, int> revisionAuthorIndexes) {
        var seenParagraphs = new HashSet<OfficeXmlNode>();
        foreach (WordElement element in elements) {
            if (element is WordParagraph wordParagraph) {
                if (!seenParagraphs.Add(wordParagraph.Node())) {
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
        var seenParagraphs = new HashSet<OfficeXmlNode>();
        foreach (WordElement element in elements) {
            if (element is WordParagraph wordParagraph) {
                if (!seenParagraphs.Add(wordParagraph.Node())) {
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
        RtfTableAlignment? tableAlignment = ToRtfTableAlignment(source.Alignment);
        OfficeXmlNode tableWidthNode = source.PropertiesNode().WChild("tblW");
        RtfTableWidthUnit? tableWidthUnit = ToRtfTableWidthUnit(tableWidthNode);
        int? tableWidth = ToRtfPreferredWidthValue(tableWidthNode);
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
        destination.HorizontalMerge = ToRtfCellMerge(source.HorizontalMerge);
        destination.VerticalMerge = ToRtfCellMerge(source.VerticalMerge);
        destination.VerticalAlignment = ToRtfCellVerticalAlignment(source.VerticalAlignment);
        destination.TextFlow = ToRtfCellTextFlow(source.TextDirection);
        OfficeXmlNode cellProperties = source.PropertiesNode();
        destination.PreferredWidthUnit = ToRtfTableWidthUnit(cellProperties.WChild("tcW"));
        destination.PreferredWidth = ToRtfPreferredWidthValue(cellProperties.WChild("tcW"));
        destination.NoWrap = !source.WrapText;
        destination.FitText = source.FitText;

        string shading = source.ShadingFillValue;
        if (!string.IsNullOrWhiteSpace(shading) && TryParseHexColor(shading, out byte red, out byte green, out byte blue)) {
            destination.BackgroundColorIndex = GetOrAddColor(document, red, green, blue);
        }

        string? foreground = WordConverterNodes.ResolveShadingForegroundColorHex(cellProperties.WChild("shd"), source.Document);
        if (!string.IsNullOrWhiteSpace(foreground) &&
            !string.Equals(foreground, "auto", StringComparison.OrdinalIgnoreCase) &&
            TryParseHexColor(foreground!, out red, out green, out blue)) {
            destination.ShadingForegroundColorIndex = GetOrAddColor(document, red, green, blue);
        }

        CopyCellShadingPattern(source.ShadingPattern, destination);
        CopyCellPadding(source, destination);
        CopyCellBorders(source, destination, document);
    }

    private static void CopyCellShadingPattern(WordShadingPattern? source, RtfTableCell destination) {
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
        CopyCellBorder(source.Borders.TopStyle, WordBorderLengths.ToSize(source.Borders.TopSize), source.Borders.TopColorValue, destination.TopBorder, document);
        CopyCellBorder(source.Borders.LeftStyle, WordBorderLengths.ToSize(source.Borders.LeftSize), source.Borders.LeftColorValue, destination.LeftBorder, document);
        CopyCellBorder(source.Borders.BottomStyle, WordBorderLengths.ToSize(source.Borders.BottomSize), source.Borders.BottomColorValue, destination.BottomBorder, document);
        CopyCellBorder(source.Borders.RightStyle, WordBorderLengths.ToSize(source.Borders.RightSize), source.Borders.RightColorValue, destination.RightBorder, document);
        CopyCellBorder(source.Borders.TopLeftToBottomRightStyle, WordBorderLengths.ToSize(source.Borders.TopLeftToBottomRightSize), source.Borders.TopLeftToBottomRightColorValue, destination.TopLeftToBottomRightBorder, document);
        CopyCellBorder(source.Borders.TopRightToBottomLeftStyle, WordBorderLengths.ToSize(source.Borders.TopRightToBottomLeftSize), source.Borders.TopRightToBottomLeftColorValue, destination.TopRightToBottomLeftBorder, document);
    }

    // RTF states border widths in twips and Word in eighths of a point.
    private static int ToRtfBorderWidth(uint wordSize) => checked((int)Math.Round(wordSize * 2.5D, MidpointRounding.AwayFromZero));

    private static uint ToWordBorderSize(int rtfWidth) => checked((uint)Math.Round(rtfWidth / 2.5D, MidpointRounding.AwayFromZero));

    private static void CopyCellBorder(WordBorderStyle? style, uint? width, string? colorHex, RtfTableCellBorder destination, RtfDocument document) {
        destination.Style = ToRtfBorderStyle(style);
        if (width.HasValue) {
            destination.Width = ToRtfBorderWidth(width.Value);
        }

        if (!string.IsNullOrWhiteSpace(colorHex) &&
            !string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase) &&
            TryParseHexColor(colorHex!, out byte red, out byte green, out byte blue)) {
            destination.ColorIndex = GetOrAddColor(document, red, green, blue);
        }
    }

    private static void CopyCellParagraphs(WordTableCell source, RtfTableCell destination, RtfDocument document, Dictionary<string, int> revisionAuthorIndexes, int tableDepth) {
        foreach (OfficeXmlNode child in source.Node().Elements()) {
            if (child.IsW("p")) {
                WordParagraph wordParagraph = source.Document.ParagraphView(child);
                RtfParagraph paragraph = destination.AddParagraph();
                CopyParagraphFormatting(wordParagraph, paragraph, document);
                AppendFormattedRuns(wordParagraph, paragraph, document, revisionAuthorIndexes);
            } else if (child.IsW("tbl")) {
                WordTable nestedWordTable = WordConverterNodes.CreateTable(source.Document, child);
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
        OfficeXmlNode indentation = table.PropertiesNode().WChild("tblInd");
        if (indentation.WAttr("type") != "dxa") {
            return null;
        }

        return WordXml.ParseInt32(indentation.WAttr("w"));
    }

    /// <summary>Recovers an RTF half-gap only from symmetric direct dxa margins and a reversible indent.</summary>
    private static int? GetWordTableCellGap(WordTable table) {
        int? left = GetWordTableLeftPadding(table);
        int? right = GetWordTableRightPadding(table);
        if (!left.HasValue || left.Value < 0 || left != right) {
            return null;
        }

        OfficeXmlNode indentation = table.PropertiesNode().WChild("tblInd");
        if (indentation.Exists &&
            (indentation.WAttr("type") != "dxa" || indentation.WAttr("w") == null)) {
            return null;
        }
        long sourceIndent = (long)(WordXml.ParseInt32(indentation.WAttr("w")) ?? 0) - GetWordTableGapIndentAdjustment(table, left.Value);
        return sourceIndent >= int.MinValue && sourceIndent <= int.MaxValue ? left : null;
    }

    /// <summary>Legacy Word placement offsets the inline grid by its leading cell margin; modern placement does not.</summary>
    private static int GetWordTableGapIndentAdjustment(WordTable table, int gap) =>
        table.Document.SourceFormat == WordFileFormat.Doc ||
        (int)table.Document.CompatibilitySettings.CompatibilityMode < (int)WordCompatibilityMode.Word2013
            ? gap : 0;

    private static int? GetWordTableLeftPadding(WordTable table) {
        OfficeXmlNode margin = table.PropertiesNode().WChild("tblCellMar").WChild("left");
        return margin.WAttr("type") == "dxa" ? WordXml.ParseInt16(margin.WAttr("w")) : null;
    }

    private static int? GetWordTableRightPadding(WordTable table) {
        OfficeXmlNode margin = table.PropertiesNode().WChild("tblCellMar").WChild("right");
        return margin.WAttr("type") == "dxa" ? WordXml.ParseInt16(margin.WAttr("w")) : null;
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
        OfficeXmlNode properties = destination.PropertiesNode();
        properties.WChild("tblCellSpacing").Remove();
        OfficeXmlNode margins = properties.GetOrAddChild(WordXml.W, "tblCellMar", WordXml.TablePropertiesOrder);
        string width = ((short)(gap ?? 0)).ToString(CultureInfo.InvariantCulture);
        OfficeXmlNode left = WordXmlEdits.ReplaceChild(margins, "left", WordXml.TableCellMarginsOrder);
        left.SetAttribute(WordXml.W, "w", width);
        left.SetAttribute(WordXml.W, "type", "dxa");
        OfficeXmlNode right = WordXmlEdits.ReplaceChild(margins, "right", WordXml.TableCellMarginsOrder);
        right.SetAttribute(WordXml.W, "w", width);
        right.SetAttribute(WordXml.W, "type", "dxa");
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
        OfficeXmlNode properties = destination.PropertiesNode();
        if (indent == 0) {
            properties.WChild("tblInd").Remove();
            return;
        }
        OfficeXmlNode indentation = WordXmlEdits.ReplaceChild(properties, "tblInd", WordXml.TablePropertiesOrder);
        indentation.SetAttribute(WordXml.W, "w", ((int)indent).ToString(CultureInfo.InvariantCulture));
        indentation.SetAttribute(WordXml.W, "type", "dxa");
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

        ApplyNativeTablePreference(destination, firstUnit, firstWidth);
    }

    // Preserve partial RTF declarations rather than inventing a missing measurement or unit.
    private static void ApplyNativeTablePreference(WordTable destination, RtfTableWidthUnit? unit, int? value) {
        destination.CheckTableProperties();
        OfficeXmlNode properties = destination.PropertiesNode();
        OfficeXmlNode existing = properties.WChild("tblW");
        OfficeXmlNode native = existing.Exists ? existing.Clone() : properties.CreateElement(WordXml.W, "tblW");
        string? type = ToWordTableWidthUnit(unit);
        if (type != null) native.SetAttribute(WordXml.W, "type", type);
        if (value.HasValue) native.SetAttribute(WordXml.W, "w", value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (existing.IsNone) {
            properties.InsertOrdered(native, WordXml.TablePropertiesOrder);
        } else {
            properties.InsertBefore(native, existing);
            existing.Remove();
        }
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
                destination.ShadingFillValue = color!;
            }
        }

        OfficeXmlNode shading;
        if (source.ShadingForegroundColorIndex.HasValue) {
            string? color = GetColorHex(document, source.ShadingForegroundColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                shading = GetOrCreateCellShading(destination);
                shading.SetAttribute(WordXml.W, "color", color!);
            }
        }

        WordShadingPattern? pattern = ToWordShadingPercent(source.ShadingPatternPercent) ?? ToWordShadingPattern(source.ShadingPattern);
        if (pattern.HasValue) {
            shading = GetOrCreateCellShading(destination);
            shading.SetAttribute(WordXml.W, "val", pattern.Value.ToOpenXmlToken());
        }

        ApplyCellPadding(source, row, destination);
        ApplyCellBorders(source, destination, document);
    }

    private static void ApplyCellPreferredWidth(RtfTableCell source, WordTableCell destination) {
        if (!source.PreferredWidthUnit.HasValue && !source.PreferredWidth.HasValue) return;
        destination.AddTableCellProperties();
        OfficeXmlNode properties = destination.PropertiesNode();
        OfficeXmlNode existing = properties.WChild("tcW");
        OfficeXmlNode native = existing.Exists ? existing.Clone() : properties.CreateElement(WordXml.W, "tcW");
        string? type = ToWordTableWidthUnit(source.PreferredWidthUnit);
        if (type != null) native.SetAttribute(WordXml.W, "type", type);
        if (source.PreferredWidth.HasValue) native.SetAttribute(WordXml.W, "w", source.PreferredWidth.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (existing.IsNone) {
            properties.InsertOrdered(native, WordXml.TableCellPropertiesOrder);
        } else {
            properties.InsertBefore(native, existing);
            existing.Remove();
        }
    }

    private static OfficeXmlNode GetOrCreateCellShading(WordTableCell cell) {
        cell.AddTableCellProperties();
        return cell.PropertiesNode().GetOrAddChild(WordXml.W, "shd", WordXml.TableCellPropertiesOrder);
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
        ApplyCellBorder(source.TopBorder, style => destination.Borders.TopStyle = style, width => destination.Borders.TopSize = WordBorderLengths.FromSize(width), color => destination.Borders.TopColorValue = color, document);
        ApplyCellBorder(source.LeftBorder, style => destination.Borders.LeftStyle = style, width => destination.Borders.LeftSize = WordBorderLengths.FromSize(width), color => destination.Borders.LeftColorValue = color, document);
        ApplyCellBorder(source.BottomBorder, style => destination.Borders.BottomStyle = style, width => destination.Borders.BottomSize = WordBorderLengths.FromSize(width), color => destination.Borders.BottomColorValue = color, document);
        ApplyCellBorder(source.RightBorder, style => destination.Borders.RightStyle = style, width => destination.Borders.RightSize = WordBorderLengths.FromSize(width), color => destination.Borders.RightColorValue = color, document);
        ApplyCellBorder(source.TopLeftToBottomRightBorder, style => destination.Borders.TopLeftToBottomRightStyle = style, width => destination.Borders.TopLeftToBottomRightSize = WordBorderLengths.FromSize(width), color => destination.Borders.TopLeftToBottomRightColorValue = color, document);
        ApplyCellBorder(source.TopRightToBottomLeftBorder, style => destination.Borders.TopRightToBottomLeftStyle = style, width => destination.Borders.TopRightToBottomLeftSize = WordBorderLengths.FromSize(width), color => destination.Borders.TopRightToBottomLeftColorValue = color, document);
    }

    private static void ApplyCellBorder(RtfTableCellBorder source, Action<WordBorderStyle?> setStyle, Action<uint?> setWidth, Action<string?> setColor, RtfDocument document) {
        if (!source.HasAnyValue) {
            return;
        }

        setStyle(ToWordBorderStyle(source.Style));
        if (source.Width.HasValue && source.Width.Value >= 0) {
            setWidth(ToWordBorderSize(source.Width.Value));
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

    private static RtfTableCellMerge ToRtfCellMerge(WordCellMerge? value) {
        if (value == WordCellMerge.Restart) return RtfTableCellMerge.First;
        if (value == WordCellMerge.Continue) return RtfTableCellMerge.Continue;
        return RtfTableCellMerge.None;
    }

    private static WordCellMerge? ToWordMergedCellValue(RtfTableCellMerge merge) {
        if (merge == RtfTableCellMerge.First) return WordCellMerge.Restart;
        if (merge == RtfTableCellMerge.Continue) return WordCellMerge.Continue;
        return null;
    }

    private static RtfTableCellVerticalAlignment? ToRtfCellVerticalAlignment(WordTableVerticalAlignment? value) {
        if (value == WordTableVerticalAlignment.Center) return RtfTableCellVerticalAlignment.Center;
        if (value == WordTableVerticalAlignment.Bottom) return RtfTableCellVerticalAlignment.Bottom;
        if (value == WordTableVerticalAlignment.Top) return RtfTableCellVerticalAlignment.Top;
        return null;
    }

    private static WordTableVerticalAlignment? ToWordCellVerticalAlignment(RtfTableCellVerticalAlignment? value) {
        if (value == RtfTableCellVerticalAlignment.Center) return WordTableVerticalAlignment.Center;
        if (value == RtfTableCellVerticalAlignment.Bottom) return WordTableVerticalAlignment.Bottom;
        if (value == RtfTableCellVerticalAlignment.Top) return WordTableVerticalAlignment.Top;
        return null;
    }

    private static RtfTableCellTextFlow? ToRtfCellTextFlow(WordTextDirection? value) {
        if (value == WordTextDirection.LeftToRightTopToBottom) return RtfTableCellTextFlow.LeftToRightTopToBottom;
        if (value == WordTextDirection.TopToBottomRightToLeft) return RtfTableCellTextFlow.TopToBottomRightToLeft;
        if (value == WordTextDirection.BottomToTopLeftToRight) return RtfTableCellTextFlow.BottomToTopLeftToRight;
        if (value == WordTextDirection.LeftToRightTopToBottomRotated) return RtfTableCellTextFlow.LeftToRightTopToBottomVertical;
        if (value == WordTextDirection.TopToBottomRightToLeftRotated) return RtfTableCellTextFlow.TopToBottomRightToLeftVertical;
        return null;
    }

    private static WordTextDirection? ToWordCellTextDirection(RtfTableCellTextFlow? value) {
        if (value == RtfTableCellTextFlow.LeftToRightTopToBottom) return WordTextDirection.LeftToRightTopToBottom;
        if (value == RtfTableCellTextFlow.TopToBottomRightToLeft) return WordTextDirection.TopToBottomRightToLeft;
        if (value == RtfTableCellTextFlow.BottomToTopLeftToRight) return WordTextDirection.BottomToTopLeftToRight;
        if (value == RtfTableCellTextFlow.LeftToRightTopToBottomVertical) return WordTextDirection.LeftToRightTopToBottomRotated;
        if (value == RtfTableCellTextFlow.TopToBottomRightToLeftVertical) return WordTextDirection.TopToBottomRightToLeftRotated;
        return null;
    }

    private static RtfTableAlignment? ToRtfTableAlignment(WordTableAlignment? value) {
        if (value == WordTableAlignment.Center) return RtfTableAlignment.Center;
        if (value == WordTableAlignment.Right) return RtfTableAlignment.Right;
        if (value == WordTableAlignment.Left) return RtfTableAlignment.Left;
        return null;
    }

    private static RtfTableWidthUnit? ToRtfTableWidthUnit(OfficeXmlNode width) {
        if (width.IsNone) return null;
        string? type = WordConverterNodes.NativeTableWidthType(width);
        if (type == null) throw new NotSupportedException("Unsupported Word preferred-width type: " + width.WAttr("type"));
        // Lexical percentages remain relative in RTF even if the native unit was omitted.
        if (width.WAttr("w")?.Trim().EndsWith("%", StringComparison.Ordinal) == true) return RtfTableWidthUnit.Percent;
        if (width.WAttr("type") == null) return null;
        if (type == "auto") return RtfTableWidthUnit.Auto;
        if (type == "dxa") return RtfTableWidthUnit.Twips;
        if (type == "pct") return RtfTableWidthUnit.Percent;
        return null;
    }

    private static int? ToRtfPreferredWidthValue(OfficeXmlNode width) {
        if (width.WAttr("w") == null) return null;
        return WordConverterNodes.TableWidthNumericValue(width) ?? WordConverterNodes.ReadTableWidth(width)?.ToNativeValue();
    }

    private static WordTableAlignment? ToWordTableAlignment(RtfTableAlignment? value) {
        if (value == RtfTableAlignment.Center) return WordTableAlignment.Center;
        if (value == RtfTableAlignment.Right) return WordTableAlignment.Right;
        if (value == RtfTableAlignment.Left) return WordTableAlignment.Left;
        return null;
    }

    // Width units are w:type tokens.
    private static string? ToWordTableWidthUnit(RtfTableWidthUnit? value) {
        if (value == RtfTableWidthUnit.Auto) return "auto";
        if (value == RtfTableWidthUnit.Twips) return "dxa";
        if (value == RtfTableWidthUnit.Percent) return "pct";
        return null;
    }

    private static RtfTableCellBorderStyle ToRtfBorderStyle(WordBorderStyle? value) {
        if (value == WordBorderStyle.Double) return RtfTableCellBorderStyle.Double;
        if (value == WordBorderStyle.Dotted) return RtfTableCellBorderStyle.Dotted;
        if (value == WordBorderStyle.Dashed) return RtfTableCellBorderStyle.Dashed;
        if (value == WordBorderStyle.Nil || value == WordBorderStyle.None) return RtfTableCellBorderStyle.None;
        if (value == WordBorderStyle.Single) return RtfTableCellBorderStyle.Single;
        return RtfTableCellBorderStyle.None;
    }

    private static WordBorderStyle? ToWordBorderStyle(RtfTableCellBorderStyle value) {
        switch (value) {
            case RtfTableCellBorderStyle.Double:
                return WordBorderStyle.Double;
            case RtfTableCellBorderStyle.Dotted:
                return WordBorderStyle.Dotted;
            case RtfTableCellBorderStyle.Dashed:
                return WordBorderStyle.Dashed;
            case RtfTableCellBorderStyle.Single:
                return WordBorderStyle.Single;
            default:
                return WordBorderStyle.Nil;
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
