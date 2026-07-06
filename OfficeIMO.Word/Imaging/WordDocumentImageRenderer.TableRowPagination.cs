using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static bool TryAddPaginatedSplitTableRow(
            WordTable table,
            IReadOnlyList<WordTableRow> rows,
            int rowIndex,
            IReadOnlyList<double> columnWidths,
            IReadOnlyList<double> rowHeights,
            int repeatingHeaderRowCount,
            double tableWidth,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers) {
            if (rowIndex < repeatingHeaderRowCount) {
                return false;
            }

            List<SplitTableCellLayout>? cells = CreateSplitTableCellLayouts(table, rows, rowIndex, columnWidths, diagnostics, listMarkers, context);
            if (cells == null || cells.Count == 0 || !cells.Any(cell => cell.HasRemainingContent)) {
                return false;
            }

            if (context.Y > context.Top && context.Y + MinimumTableRowHeightPoints > context.ContentBottom) {
                context.AdvanceColumnOrPage();
                if (context.PastTargetPage) {
                    return true;
                }

                if (!AddRepeatingTableHeaderRows(table, rows, columnWidths, rowHeights, repeatingHeaderRowCount, tableWidth, context, diagnostics, listMarkers)) {
                    return false;
                }
            }

            while (cells.Any(cell => cell.HasRemainingContent)) {
                double availableHeight = context.ContentBottom - context.Y;
                if (availableHeight < MinimumTableRowHeightPoints && context.Y > context.Top) {
                    context.AdvanceColumnOrPage();
                    if (context.PastTargetPage) {
                        return true;
                    }

                    if (!AddRepeatingTableHeaderRows(table, rows, columnWidths, rowHeights, repeatingHeaderRowCount, tableWidth, context, diagnostics, listMarkers)) {
                        return false;
                    }

                    availableHeight = context.ContentBottom - context.Y;
                }

                if (availableHeight < MinimumTableRowHeightPoints) {
                    return false;
                }

                double desiredHeight = Math.Max(MinimumTableRowHeightPoints, cells.Max(cell => cell.RemainingHeight));
                double fragmentHeight = Math.Min(availableHeight, desiredHeight);
                if (!cells.Any(cell => cell.GetContentCapacity(fragmentHeight) > 0 && cell.HasRemainingContent)) {
                    return false;
                }

                double tableLeft = ResolveTableLeft(table, context.Left, context.ContentWidth, tableWidth);
                if (context.IsTargetPage) {
                    AddSplitTableRowFragment(context.Drawing, cells, tableLeft, context.Y, fragmentHeight, diagnostics, listMarkers);
                }

                for (int i = 0; i < cells.Count; i++) {
                    cells[i].Consume(fragmentHeight);
                }

                context.Y += fragmentHeight;
                if (cells.Any(cell => cell.HasRemainingContent)) {
                    context.AdvanceColumnOrPage();
                    if (context.PastTargetPage) {
                        return true;
                    }

                    if (!AddRepeatingTableHeaderRows(table, rows, columnWidths, rowHeights, repeatingHeaderRowCount, tableWidth, context, diagnostics, listMarkers)) {
                        return false;
                    }
                }
            }

            return true;
        }

        private static List<SplitTableCellLayout>? CreateSplitTableCellLayouts(
            WordTable table,
            IReadOnlyList<WordTableRow> rows,
            int rowIndex,
            IReadOnlyList<double> columnWidths,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            WordImageFlowContext context) {
            WordTableRow row = rows[rowIndex];
            var cells = new List<SplitTableCellLayout>();
            double cellLeftOffset = 0D;
            int columnIndex = 0;
            foreach (WordTableCell cell in row.GetCells(readOnly: true)) {
                int columnSpan = Math.Max(1, cell.ColumnSpan);
                if (cell.HorizontalMerge == MergedCellValues.Continue) {
                    columnIndex += columnSpan;
                    cellLeftOffset += SumWidths(columnWidths, columnIndex - columnSpan, columnSpan);
                    continue;
                }

                if (cell.VerticalMerge == MergedCellValues.Continue) {
                    columnIndex += columnSpan;
                    cellLeftOffset += SumWidths(columnWidths, columnIndex - columnSpan, columnSpan);
                    continue;
                }

                int rowSpan = Math.Max(1, cell.RowSpan);
                List<List<WordParagraph>> paragraphRuns = CreateTableCellParagraphRuns(cell);
                bool hasListMarkers = paragraphRuns.Any(runs => CreateTableCellListMarker(runs, listMarkers).HasValue);

                WordParagraph? paragraph = paragraphRuns.Count == 0 ? null : paragraphRuns[0][0];
                double cellWidth = SumWidths(columnWidths, columnIndex, columnSpan);
                double marginLeft = ToPoints(cell.MarginLeftWidth, DefaultCellMarginPoints);
                double marginRight = ToPoints(cell.MarginRightWidth, DefaultCellMarginPoints);
                double marginTop = ToPoints(cell.MarginTopWidth, DefaultCellMarginPoints);
                double marginBottom = ToPoints(cell.MarginBottomWidth, DefaultCellMarginPoints);
                double contentWidth = Math.Max(1D, cellWidth - marginLeft - marginRight);
                List<SplitTableCellImage>? images = CreateSplitTableCellImages(cell, contentWidth, diagnostics);
                if (images == null) {
                    return null;
                }

                List<SplitTableCellNestedTable> nestedTables = CreateSplitTableCellNestedTables(cell, contentWidth);
                IReadOnlyList<SplitTableCellContentEntry> contentOrder = CreateSplitTableCellContentOrder(cell);
                A.ColorScheme? colorScheme = GetDocumentColorScheme(cell.Document);
                OfficeTextPadding padding = new OfficeTextPadding(marginLeft, marginTop, marginRight, marginBottom);
                OfficeColor fillColor = ResolveCellFillColor(table, cell, rowIndex, columnIndex, rows.Count, columnWidths.Count, colorScheme);
                OfficeBorderBox borders = ResolveCellBorders(table, cell, rowIndex, columnIndex, rowSpan, columnSpan, rows.Count, columnWidths.Count, colorScheme);
                if (ShouldSplitTableCellAsRichText(paragraphRuns, hasListMarkers)) {
                    List<OfficeRichTextRun> richRuns = CreateSplitTableCellRichRuns(paragraphRuns, colorScheme, listMarkers, context);
                    if (richRuns.Count == 0) {
                        cells.Add(SplitTableCellLayout.CreatePlain(
                            cellLeftOffset,
                            cellWidth,
                            images,
                            nestedTables,
                            contentOrder,
                            new List<string>(),
                            OfficeFontInfo.Default,
                            OfficeColor.Black,
                            OfficeTextAlignment.Left,
                            12D,
                            padding,
                            fillColor,
                            borders));
                    } else {
                        double maxFontSize = richRuns.Max(run => run.FontSize);
                        double lineHeight = Math.Max(maxFontSize * 1.25D, 12D);
                        OfficeRichTextBlockLayout richLayout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                            richRuns,
                            contentWidth,
                            double.MaxValue,
                            Math.Max(1D, lineHeight / Math.Max(1D, maxFontSize)),
                            CreateRichTextMeasure(),
                            wrap: true,
                            shrinkToFit: false,
                            minimumFontSize: Math.Min(6D, maxFontSize),
                            overflowBehavior: OfficeTextOverflowBehavior.Clip);
                        cells.Add(SplitTableCellLayout.CreateRich(
                            cellLeftOffset,
                            cellWidth,
                            images,
                            nestedTables,
                            contentOrder,
                            richLayout.Lines,
                            MapTextAlignment(paragraph?.ParagraphAlignment),
                            lineHeight,
                            padding,
                            fillColor,
                            borders));
                    }
                } else {
                    OfficeFontInfo font = paragraph == null ? OfficeFontInfo.Default : CreateFont(paragraph);
                    double lineHeight = Math.Max(font.Size * 1.25D, 12D);
                    string text = GetCellText(cell, context);
                    List<string> lines = string.IsNullOrWhiteSpace(text)
                        ? new List<string>()
                        : WrapTextIntoMeasuredLines(text, font, contentWidth);
                    cells.Add(SplitTableCellLayout.CreatePlain(
                        cellLeftOffset,
                        cellWidth,
                        images,
                        nestedTables,
                        contentOrder,
                        lines,
                        font,
                        ResolveParagraphTextColor(paragraph, colorScheme),
                        MapTextAlignment(paragraph?.ParagraphAlignment),
                        lineHeight,
                        padding,
                        fillColor,
                        borders));
                }

                columnIndex += columnSpan;
                cellLeftOffset += cellWidth;
            }

            return cells;
        }

        private static List<SplitTableCellImage>? CreateSplitTableCellImages(
            WordTableCell cell,
            double contentWidth,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var images = new List<SplitTableCellImage>();
            foreach (WordParagraph imageParagraph in cell.Elements.OfType<WordParagraph>()) {
                WordImage? image = imageParagraph.Image;
                if (image == null) {
                    continue;
                }

                if (!TryReadEmbeddedImage(image, diagnostics, out byte[] bytes, out double width, out double height)) {
                    return null;
                }

                FitImageToWidth(contentWidth, ref width, ref height);
                images.Add(new SplitTableCellImage(image, bytes, image.ContentType, width, height));
            }

            return images;
        }

        private static List<SplitTableCellNestedTable> CreateSplitTableCellNestedTables(WordTableCell cell, double contentWidth) {
            List<WordTable> nestedTables = GetDirectNestedTables(cell);
            var blocks = new List<SplitTableCellNestedTable>(nestedTables.Count);
            for (int i = 0; i < nestedTables.Count; i++) {
                double height = EstimateTableHeight(nestedTables[i], contentWidth);
                if (height > 0D) {
                    blocks.Add(new SplitTableCellNestedTable(nestedTables[i], height));
                }
            }

            return blocks;
        }

        private static IReadOnlyList<SplitTableCellContentEntry> CreateSplitTableCellContentOrder(WordTableCell cell) {
            var entries = new List<SplitTableCellContentEntry>();
            int imageIndex = 0;
            int nestedTableIndex = 0;
            bool textAdded = false;
            foreach (var child in cell._tableCell.ChildElements) {
                if (child is Paragraph paragraph) {
                    foreach (WordParagraph run in WordSection.ConvertParagraphToWordParagraphs(cell.Document, paragraph, splitPaginationMarkers: true)) {
                        if (run.IsPageBreak || run.IsColumnBreak) {
                            continue;
                        }

                        if (run.Image != null) {
                            entries.Add(SplitTableCellContentEntry.CreateImage(imageIndex++));
                        } else if (!textAdded && !string.IsNullOrEmpty(run.Text)) {
                            entries.Add(SplitTableCellContentEntry.CreateText());
                            textAdded = true;
                        }
                    }
                } else if (child is Table) {
                    entries.Add(SplitTableCellContentEntry.CreateNestedTable(nestedTableIndex++));
                }
            }

            return entries;
        }

        private static bool ShouldSplitTableCellAsRichText(IReadOnlyList<IReadOnlyList<WordParagraph>> paragraphRuns, bool hasListMarkers) {
            if (paragraphRuns.Count > 1 || hasListMarkers) {
                return true;
            }

            return paragraphRuns.Count == 1 && (paragraphRuns[0].Count > 1 || HasRunHighlight(paragraphRuns[0][0]));
        }

        private static List<OfficeRichTextRun> CreateSplitTableCellRichRuns(
            IReadOnlyList<IReadOnlyList<WordParagraph>> paragraphRuns,
            A.ColorScheme? colorScheme,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            WordImageFlowContext? context = null) {
            var richRuns = new List<OfficeRichTextRun>();
            for (int paragraphIndex = 0; paragraphIndex < paragraphRuns.Count; paragraphIndex++) {
                IReadOnlyList<WordParagraph> runs = paragraphRuns[paragraphIndex];
                if (runs.Count == 0) {
                    continue;
                }

                if (richRuns.Count > 0) {
                    richRuns.Add(CreateRichTextRun(runs[0], colorScheme, Environment.NewLine));
                }

                WordImageListMarker? listMarker = CreateTableCellListMarker(runs, listMarkers);
                if (listMarker.HasValue) {
                    WordImageListMarker marker = listMarker.Value;
                    richRuns.Add(new OfficeRichTextRun(
                        marker.Marker + " ",
                        marker.Font.Size,
                        marker.Color,
                        marker.Font.IsBold,
                        marker.Font.IsItalic,
                        marker.Font.IsUnderline,
                        marker.Font.FamilyName,
                        marker.Font.IsStrikethrough));
                }

                richRuns.AddRange(CreateRichTextRuns(runs, colorScheme, context));
            }

            return richRuns;
        }

        private static void AddSplitTableRowFragment(
            OfficeDrawing drawing,
            IReadOnlyList<SplitTableCellLayout> cells,
            double tableLeft,
            double top,
            double height,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers) {
            for (int i = 0; i < cells.Count; i++) {
                SplitTableCellLayout cell = cells[i];
                double left = tableLeft + cell.LeftOffset;
                cell.AddFrameAndText(drawing, left, top, height, diagnostics, listMarkers);
            }
        }

        private sealed class SplitTableCellLayout {
            private readonly double _leftOffset;
            private readonly double _width;
            private readonly IReadOnlyList<SplitTableCellImage> _images;
            private readonly IReadOnlyList<SplitTableCellNestedTable> _nestedTables;
            private readonly IReadOnlyList<SplitTableCellContentEntry> _contentOrder;
            private readonly List<string>? _plainLines;
            private readonly IReadOnlyList<OfficeRichTextLine>? _richLines;
            private readonly OfficeFontInfo? _font;
            private readonly OfficeColor _textColor;
            private readonly OfficeTextAlignment _alignment;
            private readonly double _lineHeight;
            private readonly OfficeTextPadding _padding;
            private readonly OfficeColor _fillColor;
            private readonly OfficeBorderBox _borders;
            private int _imageIndex;
            private int _nestedTableIndex;
            private int _lineIndex;

            private SplitTableCellLayout(
                double leftOffset,
                double width,
                IReadOnlyList<SplitTableCellImage> images,
                IReadOnlyList<SplitTableCellNestedTable> nestedTables,
                IReadOnlyList<SplitTableCellContentEntry> contentOrder,
                List<string>? plainLines,
                IReadOnlyList<OfficeRichTextLine>? richLines,
                OfficeFontInfo? font,
                OfficeColor textColor,
                OfficeTextAlignment alignment,
                double lineHeight,
                OfficeTextPadding padding,
                OfficeColor fillColor,
                OfficeBorderBox borders) {
                _leftOffset = leftOffset;
                _width = width;
                _images = images;
                _nestedTables = nestedTables;
                _contentOrder = contentOrder;
                _plainLines = plainLines;
                _richLines = richLines;
                _font = font;
                _textColor = textColor;
                _alignment = alignment;
                _lineHeight = lineHeight;
                _padding = padding;
                _fillColor = fillColor;
                _borders = borders;
            }

            internal static SplitTableCellLayout CreatePlain(
                double leftOffset,
                double width,
                IReadOnlyList<SplitTableCellImage> images,
                IReadOnlyList<SplitTableCellNestedTable> nestedTables,
                IReadOnlyList<SplitTableCellContentEntry> contentOrder,
                List<string> lines,
                OfficeFontInfo font,
                OfficeColor textColor,
                OfficeTextAlignment alignment,
                double lineHeight,
                OfficeTextPadding padding,
                OfficeColor fillColor,
                OfficeBorderBox borders) =>
                new SplitTableCellLayout(leftOffset, width, images, nestedTables, contentOrder, lines, null, font, textColor, alignment, lineHeight, padding, fillColor, borders);

            internal static SplitTableCellLayout CreateRich(
                double leftOffset,
                double width,
                IReadOnlyList<SplitTableCellImage> images,
                IReadOnlyList<SplitTableCellNestedTable> nestedTables,
                IReadOnlyList<SplitTableCellContentEntry> contentOrder,
                IReadOnlyList<OfficeRichTextLine> lines,
                OfficeTextAlignment alignment,
                double lineHeight,
                OfficeTextPadding padding,
                OfficeColor fillColor,
                OfficeBorderBox borders) =>
                new SplitTableCellLayout(leftOffset, width, images, nestedTables, contentOrder, null, lines, null, OfficeColor.Black, alignment, lineHeight, padding, fillColor, borders);

            internal bool HasRemainingContent => _imageIndex < ImageCount || _nestedTableIndex < NestedTableCount || _lineIndex < LineCount;

            internal double LeftOffset => _leftOffset;

            internal double RemainingHeight {
                get {
                    double height = _padding.Top + _padding.Bottom;
                    for (int i = _imageIndex; i < ImageCount; i++) {
                        height += _images[i].Height;
                        if (i < ImageCount - 1 || _nestedTableIndex < NestedTableCount || _lineIndex < LineCount) {
                            height += ParagraphGapPoints;
                        }
                    }

                    for (int i = _nestedTableIndex; i < NestedTableCount; i++) {
                        height += _nestedTables[i].Height;
                        if (i < NestedTableCount - 1 || _lineIndex < LineCount) {
                            height += ParagraphGapPoints;
                        }
                    }

                    for (int i = _lineIndex; i < LineCount; i++) {
                        height += GetLineHeight(i);
                    }

                    return height;
                }
            }

            internal int GetContentCapacity(double fragmentHeight) {
                SplitTableCellFragment fragment = CreateFragment(fragmentHeight);
                return fragment.ImageCount + fragment.NestedTableCount + fragment.LineCount;
            }

            internal int GetLineCapacity(double fragmentHeight) {
                SplitTableCellFragment fragment = CreateFragment(fragmentHeight);
                return fragment.LineCount;
            }

            private SplitTableCellFragment CreateFragment(double fragmentHeight) {
                double textHeight = Math.Max(0D, fragmentHeight - _padding.Top - _padding.Bottom);
                double used = 0D;
                int imageCount = 0;
                int nestedTableCount = 0;
                int lineCount = 0;
                for (int i = _imageIndex; i < ImageCount; i++) {
                    double imageHeight = _images[i].Height;
                    if (i < ImageCount - 1 || _nestedTableIndex < NestedTableCount || _lineIndex < LineCount) {
                        imageHeight += ParagraphGapPoints;
                    }

                    if (imageCount > 0 && used + imageHeight > textHeight + 0.01D) {
                        return new SplitTableCellFragment(imageCount, nestedTableCount, lineCount);
                    }

                    if (imageCount == 0 && imageHeight > textHeight + 0.01D) {
                        return new SplitTableCellFragment(0, 0, 0);
                    }

                    used += imageHeight;
                    imageCount++;
                }

                for (int i = _nestedTableIndex; i < NestedTableCount; i++) {
                    double nestedTableHeight = _nestedTables[i].Height;
                    if (i < NestedTableCount - 1 || _lineIndex < LineCount) {
                        nestedTableHeight += ParagraphGapPoints;
                    }

                    if ((imageCount > 0 || nestedTableCount > 0) && used + nestedTableHeight > textHeight + 0.01D) {
                        return new SplitTableCellFragment(imageCount, nestedTableCount, lineCount);
                    }

                    if (imageCount == 0 && nestedTableCount == 0 && nestedTableHeight > textHeight + 0.01D) {
                        return new SplitTableCellFragment(0, 0, 0);
                    }

                    used += nestedTableHeight;
                    nestedTableCount++;
                }

                for (int i = _lineIndex; i < LineCount; i++) {
                    double lineHeight = GetLineHeight(i);
                    if ((imageCount > 0 || nestedTableCount > 0 || lineCount > 0) && used + lineHeight > textHeight + 0.01D) {
                        break;
                    }

                    if (imageCount == 0 && nestedTableCount == 0 && lineCount == 0 && lineHeight > textHeight + 0.01D) {
                        break;
                    }

                    used += lineHeight;
                    lineCount++;
                }

                return new SplitTableCellFragment(imageCount, nestedTableCount, lineCount);
            }

            internal void Consume(double fragmentHeight) {
                SplitTableCellFragment fragment = CreateFragment(fragmentHeight);
                _imageIndex += Math.Min(fragment.ImageCount, Math.Max(0, ImageCount - _imageIndex));
                _nestedTableIndex += Math.Min(fragment.NestedTableCount, Math.Max(0, NestedTableCount - _nestedTableIndex));
                _lineIndex += Math.Min(fragment.LineCount, Math.Max(0, LineCount - _lineIndex));
            }

            internal void AddFrameAndText(
                OfficeDrawing drawing,
                double left,
                double top,
                double height,
                List<OfficeImageExportDiagnostic> diagnostics,
                IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers) {
                drawing.AddBorderBox(left, top, _width, height, _fillColor, _borders);
                SplitTableCellFragment fragment = CreateFragment(height);
                if (fragment.ImageCount <= 0 && fragment.NestedTableCount <= 0 && fragment.LineCount <= 0) {
                    return;
                }

                double contentLeft = left + _padding.Left;
                double contentTop = top + _padding.Top;
                double contentWidth = Math.Max(1D, _width - _padding.Left - _padding.Right);
                int lineCount = Math.Min(fragment.LineCount, Math.Max(0, LineCount - _lineIndex));
                int imageStart = _imageIndex;
                int imageEnd = imageStart + fragment.ImageCount;
                int nestedTableStart = _nestedTableIndex;
                int nestedTableEnd = nestedTableStart + fragment.NestedTableCount;
                var drawnImages = new HashSet<int>();
                var drawnNestedTables = new HashSet<int>();
                bool drewText = false;
                int drawnBlocks = 0;
                int totalBlocks = fragment.ImageCount + fragment.NestedTableCount + (lineCount > 0 ? 1 : 0);

                void AddGapIfNeeded() {
                    drawnBlocks++;
                    if (drawnBlocks < totalBlocks) {
                        contentTop += ParagraphGapPoints;
                    }
                }

                bool DrawImageBlock(int imageIndex) {
                    if (imageIndex < imageStart || imageIndex >= imageEnd || !drawnImages.Add(imageIndex)) {
                        return false;
                    }

                    SplitTableCellImage image = _images[imageIndex];
                    drawing.AddImage(
                        image.Bytes,
                        image.ContentType,
                        CreateImageProjection(image.Source, contentLeft, contentTop, image.Width, image.Height),
                        DescribeImage(image.Source));
                    contentTop += image.Height;
                    AddGapIfNeeded();
                    return true;
                }

                bool DrawNestedTableBlock(int nestedTableIndex) {
                    if (nestedTableIndex < nestedTableStart || nestedTableIndex >= nestedTableEnd || !drawnNestedTables.Add(nestedTableIndex)) {
                        return false;
                    }

                    SplitTableCellNestedTable nestedTable = _nestedTables[nestedTableIndex];
                    WordImageFlowContext nestedContext = CreateFlowContext(
                        drawing,
                        contentLeft,
                        contentTop,
                        contentWidth,
                        contentTop + nestedTable.Height,
                        "unsupported-word-nested-table-overflow",
                        "Skipped a nested Word table inside a split table row because it does not fit within the row fragment.");
                    AddTable(nestedTable.Table, nestedContext, diagnostics, listMarkers, allowNestedTable: true);
                    contentTop += nestedTable.Height;
                    AddGapIfNeeded();
                    return true;
                }

                bool DrawTextBlock() {
                    if (drewText || lineCount <= 0) {
                        return false;
                    }

                    double textTop = contentTop > top + _padding.Top + 0.000001D ? contentTop - _padding.Top : top;
                    double textHeight = Math.Max(1D, top + height - textTop);
                    if (_richLines != null) {
                        drawing.AddRichText(
                            CreateRichTextRunsFromLines(_richLines, _lineIndex, lineCount),
                            left,
                            textTop,
                            _width,
                            textHeight,
                            _alignment,
                            _lineHeight,
                            OfficeTextVerticalAlignment.Top,
                            wrapText: true,
                            padding: _padding);
                    } else {
                        string text = string.Join(Environment.NewLine, _plainLines!.GetRange(_lineIndex, lineCount).Select(line => line.TrimEnd(' ', '\t')));
                        drawing.AddText(
                            text,
                            left,
                            textTop,
                            _width,
                            textHeight,
                            _font,
                            _textColor,
                            _alignment,
                            _lineHeight,
                            OfficeTextVerticalAlignment.Top,
                            wrapText: true,
                            padding: _padding);
                    }

                    drewText = true;
                    AddGapIfNeeded();
                    return true;
                }

                for (int i = 0; i < _contentOrder.Count; i++) {
                    SplitTableCellContentEntry entry = _contentOrder[i];
                    if (entry.Kind == SplitTableCellContentKind.Image) {
                        DrawImageBlock(entry.Index);
                    } else if (entry.Kind == SplitTableCellContentKind.NestedTable) {
                        DrawNestedTableBlock(entry.Index);
                    } else {
                        DrawTextBlock();
                    }
                }

                for (int i = imageStart; i < imageEnd; i++) {
                    DrawImageBlock(i);
                }

                for (int i = nestedTableStart; i < nestedTableEnd; i++) {
                    DrawNestedTableBlock(i);
                }

                DrawTextBlock();
            }

            private int ImageCount => _images.Count;

            private int NestedTableCount => _nestedTables.Count;

            private int LineCount => _richLines?.Count ?? _plainLines?.Count ?? 0;

            private double GetLineHeight(int index) =>
                _richLines != null ? ResolveRichTextSliceLineHeight(_richLines[index], _lineHeight) : _lineHeight;
        }

        private enum SplitTableCellContentKind {
            Image,
            NestedTable,
            Text
        }

        private readonly struct SplitTableCellContentEntry {
            private SplitTableCellContentEntry(SplitTableCellContentKind kind, int index) {
                Kind = kind;
                Index = index;
            }

            internal SplitTableCellContentKind Kind { get; }

            internal int Index { get; }

            internal static SplitTableCellContentEntry CreateImage(int index) =>
                new(SplitTableCellContentKind.Image, index);

            internal static SplitTableCellContentEntry CreateNestedTable(int index) =>
                new(SplitTableCellContentKind.NestedTable, index);

            internal static SplitTableCellContentEntry CreateText() =>
                new(SplitTableCellContentKind.Text, -1);
        }

        private readonly struct SplitTableCellImage {
            internal SplitTableCellImage(WordImage source, byte[] bytes, string contentType, double width, double height) {
                Source = source;
                Bytes = bytes;
                ContentType = contentType;
                Width = width;
                Height = height;
            }

            internal WordImage Source { get; }

            internal byte[] Bytes { get; }

            internal string ContentType { get; }

            internal double Width { get; }

            internal double Height { get; }
        }

        private readonly struct SplitTableCellNestedTable {
            internal SplitTableCellNestedTable(WordTable table, double height) {
                Table = table;
                Height = height;
            }

            internal WordTable Table { get; }

            internal double Height { get; }
        }

        private readonly struct SplitTableCellFragment {
            internal SplitTableCellFragment(int imageCount, int nestedTableCount, int lineCount) {
                ImageCount = imageCount;
                NestedTableCount = nestedTableCount;
                LineCount = lineCount;
            }

            internal int ImageCount { get; }

            internal int NestedTableCount { get; }

            internal int LineCount { get; }
        }
    }
}
