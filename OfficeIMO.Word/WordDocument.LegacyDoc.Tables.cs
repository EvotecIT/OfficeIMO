using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static void AddLegacyDocTable(WordSection section, LegacyDocTableBlock tableBlock, LegacyDocStyleSheet styleSheet, LegacyDocNoteProjection notes) {
            AddLegacyDocTableCore((rows, columns) => section.AddTable(rows, columns, WordTableStyle.TableNormal), tableBlock, styleSheet, notes);
        }

        private static void AddLegacyDocTableCore(Func<int, int, WordTable> createTable, LegacyDocTableBlock tableBlock, LegacyDocStyleSheet styleSheet, LegacyDocNoteProjection notes, bool projectNestedTables = true) {
            int rowCount = tableBlock.Rows.Count;
            int columnCount = tableBlock.Rows.Count == 0
                ? 0
                : tableBlock.Rows.Max(row => row.Cells.Count);
            if (rowCount == 0 || columnCount == 0) {
                return;
            }

            WordTable table = createTable(rowCount, columnCount);
            ApplyLegacyDocTableBorderDefaults(table, tableBlock.Rows.Select(row =>
                row.TableBorders.WithDefaults(styleSheet.ResolveTableBorders(row.TableStyleIndex))));
            LegacyDocTableAlignment? tableAlignment = tableBlock.Rows
                .Select(row => row.TableAlignment)
                .FirstOrDefault(alignment => alignment.HasValue);
            if (tableAlignment != null) {
                ApplyLegacyDocTableAlignment(table, tableAlignment.Value);
            }

            int? tableLeftIndentTwips = tableBlock.Rows
                .Select(row => row.TableLeftIndentTwips)
                .FirstOrDefault(indent => indent.HasValue);
            if (tableLeftIndentTwips != null) {
                ApplyLegacyDocTableIndentation(table, tableLeftIndentTwips.Value);
            }

            LegacyDocTablePreferredWidth? tablePreferredWidth = tableBlock.Rows
                .Select(row => row.TablePreferredWidth)
                .FirstOrDefault(width => width.HasValue);
            if (tablePreferredWidth != null) {
                ApplyLegacyDocTablePreferredWidth(table, tablePreferredWidth.Value);
            }

            bool? tableAutofit = tableBlock.Rows
                .Select(row => row.TableAutofit)
                .FirstOrDefault(autofit => autofit.HasValue);
            // Binary DOC defaults to fixed columns when sprmTFAutofit is absent;
            // the projected DOCX table must not acquire DOCX's AutoFit default.
            table.LayoutMode = tableAutofit == true ? WordTableLayoutMode.AutoFit : WordTableLayoutMode.Fixed;

            int? tableCellSpacingTwips = tableBlock.Rows
                .Select(row => row.DefaultCellSpacingTwips)
                .FirstOrDefault(spacing => spacing.HasValue);
            if (tableCellSpacingTwips != null) {
                table.StyleDetails!.CellSpacing = checked((short)tableCellSpacingTwips.Value);
            }

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                LegacyDocTableRow sourceRow = tableBlock.Rows[rowIndex];
                ApplyLegacyDocTableRowFormatting(table.Rows[rowIndex], sourceRow);
                for (int columnIndex = 0; columnIndex < sourceRow.Cells.Count && columnIndex < columnCount; columnIndex++) {
                    AddLegacyDocTableCell(table.Rows[rowIndex].Cells[columnIndex], sourceRow.Cells[columnIndex], styleSheet, notes, projectNestedTables);
                    if (columnIndex < sourceRow.CellWidthsTwips.Count) {
                        table.Rows[rowIndex].Cells[columnIndex].WidthType = WordTableWidthUnit.Dxa;
                        table.Rows[rowIndex].Cells[columnIndex].Width = sourceRow.CellWidthsTwips[columnIndex];
                    }

                    if (columnIndex < sourceRow.CellHorizontalMerges.Count) {
                        ApplyLegacyDocTableCellHorizontalMerge(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellHorizontalMerges[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellVerticalMerges.Count) {
                        ApplyLegacyDocTableCellVerticalMerge(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellVerticalMerges[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellVerticalAlignments.Count) {
                        ApplyLegacyDocTableCellVerticalAlignment(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellVerticalAlignments[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellTextDirections.Count) {
                        ApplyLegacyDocTableCellTextDirection(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellTextDirections[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellFitTexts.Count && sourceRow.CellFitTexts[columnIndex]) {
                        table.Rows[rowIndex].Cells[columnIndex].FitText = true;
                    }

                    if (columnIndex < sourceRow.CellNoWraps.Count && sourceRow.CellNoWraps[columnIndex]) {
                        table.Rows[rowIndex].Cells[columnIndex].WrapText = false;
                    }

                    if (columnIndex < sourceRow.CellHideMarks.Count && sourceRow.CellHideMarks[columnIndex]) {
                        table.Rows[rowIndex].Cells[columnIndex].HideMark = true;
                    }

                    if (columnIndex < sourceRow.CellMargins.Count) {
                        ApplyLegacyDocTableCellMargins(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellMargins[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellShadings.Count) {
                        ApplyLegacyDocTableCellShading(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellShadings[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellBorders.Count) {
                        ApplyLegacyDocTableCellBorders(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellBorders[columnIndex]);
                    }
                }
            }

            AddLegacyDocTableRowBoundaryBookmarks(table, tableBlock);
            AddLegacyDocTableBlockBookmarks(table, tableBlock);
        }

        private static void AddLegacyDocTableRowBoundaryBookmarks(WordTable table, LegacyDocTableBlock tableBlock) {
            int rowCount = Math.Min(table.Rows.Count, tableBlock.Rows.Count);
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                LegacyDocTableRow sourceRow = tableBlock.Rows[rowIndex];
                if (sourceRow.BookmarksBefore.Count == 0) {
                    continue;
                }

                TableRow row = table.Rows[rowIndex]._tableRow;
                foreach (LegacyDocBookmark bookmark in sourceRow.BookmarksBefore.OrderBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                    table._table.InsertBefore(new BookmarkStart { Id = bookmark.ProjectionId, Name = bookmark.Name }, row);
                    table._table.InsertBefore(new BookmarkEnd { Id = bookmark.ProjectionId }, row);
                }
            }
        }

        private static void AddLegacyDocTableBlockBookmarks(WordTable table, LegacyDocTableBlock tableBlock) {
            if (tableBlock.Bookmarks.Count == 0 || table._table.Parent is not OpenXmlCompositeElement parent) {
                return;
            }

            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => bookmark.IsZeroLength && bookmark.StartCharacter == tableBlock.StartCharacter)
                .OrderBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                parent.InsertBefore(new BookmarkStart { Id = bookmark.ProjectionId, Name = bookmark.Name }, table._table);
                parent.InsertBefore(new BookmarkEnd { Id = bookmark.ProjectionId }, table._table);
            }

            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => !bookmark.IsZeroLength && bookmark.StartCharacter == tableBlock.StartCharacter)
                .OrderByDescending(bookmark => bookmark.EndCharacter)
                .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                parent.InsertBefore(new BookmarkStart { Id = bookmark.ProjectionId, Name = bookmark.Name }, table._table);
            }

            OpenXmlElement afterAnchor = table._table;
            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => bookmark.EndCharacter == tableBlock.EndCharacter && !bookmark.IsZeroLength)
                .OrderByDescending(bookmark => bookmark.StartCharacter)
                .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                afterAnchor = parent.InsertAfter(new BookmarkEnd { Id = bookmark.ProjectionId }, afterAnchor)!;
            }

            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => bookmark.IsZeroLength && bookmark.StartCharacter == tableBlock.EndCharacter)
                .OrderBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                afterAnchor = parent.InsertAfter(new BookmarkStart { Id = bookmark.ProjectionId, Name = bookmark.Name }, afterAnchor)!;
                afterAnchor = parent.InsertAfter(new BookmarkEnd { Id = bookmark.ProjectionId }, afterAnchor)!;
            }
        }

        private static void ApplyLegacyDocTablePreferredWidth(WordTable table, LegacyDocTablePreferredWidth preferredWidth) {
            switch (preferredWidth.Unit) {
                case LegacyDocTablePreferredWidthUnit.Auto:
                    table.WidthType = WordTableWidthUnit.Auto;
                    table.Width = 0;
                    break;
                case LegacyDocTablePreferredWidthUnit.Percent:
                    table.WidthType = WordTableWidthUnit.Pct;
                    table.Width = preferredWidth.Value;
                    break;
                case LegacyDocTablePreferredWidthUnit.Dxa:
                    table.WidthType = WordTableWidthUnit.Dxa;
                    table.Width = preferredWidth.Value;
                    break;
            }
        }

        private static void ApplyLegacyDocTableAlignment(WordTable table, LegacyDocTableAlignment tableAlignment) {
            switch (tableAlignment) {
                case LegacyDocTableAlignment.Left:
                    table.Alignment = WordTableAlignment.Left;
                    break;
                case LegacyDocTableAlignment.Center:
                    table.Alignment = WordTableAlignment.Center;
                    break;
                case LegacyDocTableAlignment.Right:
                    table.Alignment = WordTableAlignment.Right;
                    break;
            }
        }

        private static void ApplyLegacyDocTableIndentation(WordTable table, int leftIndentTwips) {
            table.CheckTableProperties();
            table._tableProperties!.TableIndentation = new TableIndentation {
                Width = leftIndentTwips,
                Type = TableWidthUnitValues.Dxa
            };
        }

        private static void ApplyLegacyDocTableCellMargins(WordTableCell cell, LegacyDocTableCellMargins margins) {
            if (margins.TopTwips != null) {
                cell.MarginTopWidth = checked((short)margins.TopTwips.Value);
            }

            if (margins.RightTwips != null) {
                cell.MarginRightWidth = checked((short)margins.RightTwips.Value);
            }

            if (margins.BottomTwips != null) {
                cell.MarginBottomWidth = checked((short)margins.BottomTwips.Value);
            }

            if (margins.LeftTwips != null) {
                cell.MarginLeftWidth = checked((short)margins.LeftTwips.Value);
            }
        }

        private static void ApplyLegacyDocTableCellShading(WordTableCell cell, LegacyDocTableCellShading shading) {
            if (!string.IsNullOrEmpty(shading.FillColorHex)) {
                cell.ShadingFillColorHex = shading.FillColorHex!;
            }
        }

        private static void ApplyLegacyDocTableCellBorders(WordTableCell cell, LegacyDocTableCellBorders borders) {
            ApplyLegacyDocTableCellBorder(
                borders.Top,
                style => cell.Borders.TopStyle = style.ToOfficeEnum(),
                color => cell.Borders.TopColorHex = color,
                size => cell.Borders.TopSize = size,
                space => cell.Borders.TopSpace = space);
            ApplyLegacyDocTableCellBorder(
                borders.Left,
                style => cell.Borders.LeftStyle = style.ToOfficeEnum(),
                color => cell.Borders.LeftColorHex = color,
                size => cell.Borders.LeftSize = size,
                space => cell.Borders.LeftSpace = space);
            ApplyLegacyDocTableCellBorder(
                borders.Bottom,
                style => cell.Borders.BottomStyle = style.ToOfficeEnum(),
                color => cell.Borders.BottomColorHex = color,
                size => cell.Borders.BottomSize = size,
                space => cell.Borders.BottomSpace = space);
            ApplyLegacyDocTableCellBorder(
                borders.Right,
                style => cell.Borders.RightStyle = style.ToOfficeEnum(),
                color => cell.Borders.RightColorHex = color,
                size => cell.Borders.RightSize = size,
                space => cell.Borders.RightSpace = space);
        }

        private static void ApplyLegacyDocTableCellBorder(
            LegacyDocTableCellBorder border,
            Action<BorderValues> setStyle,
            Action<string?> setColor,
            Action<uint?> setSize,
            Action<uint?> setSpace) {
            BorderValues? style = MapLegacyDocTableCellBorderStyle(border.Style);
            if (style == null) {
                return;
            }

            setStyle(style.Value);
            // A specified edge replaces the earlier row default as a complete value.
            setColor(border.ColorHex);
            setSize(border.SizeEighthPoints > 0 ? (uint?)border.SizeEighthPoints : null);
            setSpace(border.SpacePoints > 0 ? (uint?)border.SpacePoints : null);
        }

        private static BorderValues? MapLegacyDocTableCellBorderStyle(LegacyDocTableCellBorderStyle style) {
            switch (style) {
                case LegacyDocTableCellBorderStyle.ExplicitNone:
                    return BorderValues.Nil;
                case LegacyDocTableCellBorderStyle.Single:
                    return BorderValues.Single;
                case LegacyDocTableCellBorderStyle.Double:
                    return BorderValues.Double;
                case LegacyDocTableCellBorderStyle.Dotted:
                    return BorderValues.Dotted;
                case LegacyDocTableCellBorderStyle.Dashed:
                    return BorderValues.Dashed;
                default:
                    return null;
            }
        }

        private static void ApplyLegacyDocTableCellHorizontalMerge(WordTableCell cell, LegacyDocTableCellHorizontalMerge horizontalMerge) {
            switch (horizontalMerge) {
                case LegacyDocTableCellHorizontalMerge.Restart:
                    cell.HorizontalMerge = WordCellMerge.Restart;
                    break;
                case LegacyDocTableCellHorizontalMerge.Continue:
                    cell.HorizontalMerge = WordCellMerge.Continue;
                    break;
            }
        }

        private static void ApplyLegacyDocTableCellVerticalMerge(WordTableCell cell, LegacyDocTableCellVerticalMerge verticalMerge) {
            switch (verticalMerge) {
                case LegacyDocTableCellVerticalMerge.Restart:
                    cell.VerticalMerge = WordCellMerge.Restart;
                    break;
                case LegacyDocTableCellVerticalMerge.Continue:
                    cell.VerticalMerge = WordCellMerge.Continue;
                    break;
            }
        }

        private static void ApplyLegacyDocTableCellVerticalAlignment(WordTableCell cell, LegacyDocTableCellVerticalAlignment verticalAlignment) {
            switch (verticalAlignment) {
                case LegacyDocTableCellVerticalAlignment.Center:
                    cell.VerticalAlignment = WordTableVerticalAlignment.Center;
                    break;
                case LegacyDocTableCellVerticalAlignment.Bottom:
                    cell.VerticalAlignment = WordTableVerticalAlignment.Bottom;
                    break;
            }
        }

        private static void ApplyLegacyDocTableCellTextDirection(WordTableCell cell, LegacyDocTableCellTextDirection textDirection) {
            switch (textDirection) {
                case LegacyDocTableCellTextDirection.TopToBottomRightToLeft:
                    cell.TextDirection = WordTextDirection.TopToBottomRightToLeft;
                    break;
                case LegacyDocTableCellTextDirection.BottomToTopLeftToRight:
                    cell.TextDirection = WordTextDirection.BottomToTopLeftToRight;
                    break;
                case LegacyDocTableCellTextDirection.LeftToRightTopToBottomRotated:
                    cell.TextDirection = WordTextDirection.LeftToRightTopToBottomRotated;
                    break;
                case LegacyDocTableCellTextDirection.TopToBottomRightToLeftRotated:
                    cell.TextDirection = WordTextDirection.TopToBottomRightToLeftRotated;
                    break;
            }
        }

        private static void ApplyLegacyDocTableRowFormatting(WordTableRow row, LegacyDocTableRow sourceRow) {
            if (sourceRow.RowHeightTwips != null) {
                row.AddTableRowProperties();
                TableRowProperties rowProperties = row._tableRow.TableRowProperties!;
                TableRowHeight? rowHeight = rowProperties.GetFirstChild<TableRowHeight>();
                if (rowHeight == null) {
                    rowHeight = new TableRowHeight();
                    rowProperties.InsertAt(rowHeight, 0);
                }

                rowHeight.Val = (uint)sourceRow.RowHeightTwips.Value;
                rowHeight.HeightType = sourceRow.RowHeightIsExact
                    ? HeightRuleValues.Exact
                    : HeightRuleValues.AtLeast;
            }

            if (sourceRow.RowCantSplit == true) {
                row.AllowRowToBreakAcrossPages = false;
            }

            if (sourceRow.RowIsHeader == true) {
                row.RepeatHeaderRowAtTheTopOfEachPage = true;
            }
        }

        private static void AddLegacyDocTableCell(WordTableCell cell, LegacyDocTableCell sourceCell, LegacyDocStyleSheet styleSheet, LegacyDocNoteProjection notes, bool projectNestedTables = true) {
            var pendingBookmarks = new List<LegacyDocBookmark>();
            int pendingBookmarkStartCharacter = int.MaxValue;
            int pendingBookmarkEndCharacter = int.MinValue;
            bool emittedParagraph = false;
            for (int index = 0; index < sourceCell.Paragraphs.Count; index++) {
                LegacyDocTableCellParagraph sourceParagraph = sourceCell.Paragraphs[index];
                if (string.IsNullOrEmpty(sourceParagraph.Text)
                    && sourceParagraph.Bookmarks.Count > 0
                    && index + 1 < sourceCell.Paragraphs.Count) {
                    AddPendingTableCellBookmarks(sourceParagraph.Bookmarks, ref pendingBookmarkStartCharacter, ref pendingBookmarkEndCharacter, pendingBookmarks);
                    continue;
                }

                if (projectNestedTables && IsLegacyDocNestedTableParagraph(sourceParagraph)) {
                    index = AddLegacyDocNestedTable(
                        cell,
                        sourceCell.Paragraphs,
                        index,
                        styleSheet,
                        notes,
                        pendingBookmarks);
                    emittedParagraph = true;
                    pendingBookmarks.Clear();
                    pendingBookmarkStartCharacter = int.MaxValue;
                    pendingBookmarkEndCharacter = int.MinValue;
                    continue;
                }

                AddLegacyDocTableCellParagraph(
                    cell,
                    sourceParagraph,
                    styleSheet,
                    notes,
                    removeExistingParagraphs: !emittedParagraph,
                    pendingBookmarks,
                    pendingBookmarkStartCharacter,
                    pendingBookmarkEndCharacter);
                emittedParagraph = true;
                pendingBookmarks.Clear();
                pendingBookmarkStartCharacter = int.MaxValue;
                pendingBookmarkEndCharacter = int.MinValue;
            }
        }

        private static void AddPendingTableCellBookmarks(IReadOnlyList<LegacyDocBookmark> bookmarks, ref int startCharacter, ref int endCharacter, List<LegacyDocBookmark> pendingBookmarks) {
            foreach (LegacyDocBookmark bookmark in bookmarks) {
                pendingBookmarks.Add(bookmark);
                startCharacter = Math.Min(startCharacter, bookmark.StartCharacter);
                endCharacter = Math.Max(endCharacter, bookmark.EndCharacter);
            }
        }

        private static void AddLegacyDocTableCellParagraph(
            WordTableCell cell,
            LegacyDocTableCellParagraph sourceParagraph,
            LegacyDocStyleSheet styleSheet,
            LegacyDocNoteProjection notes,
            bool removeExistingParagraphs,
            IReadOnlyList<LegacyDocBookmark>? pendingBookmarks = null,
            int pendingBookmarkStartCharacter = int.MaxValue,
            int pendingBookmarkEndCharacter = int.MinValue) {
            IReadOnlyList<LegacyDocBookmark> paragraphBookmarks = MergeLegacyDocTableCellBookmarks(sourceParagraph.Bookmarks, pendingBookmarks);
            int paragraphStartCharacter = pendingBookmarks != null && pendingBookmarks.Count > 0
                ? Math.Min(sourceParagraph.StartCharacter, pendingBookmarkStartCharacter)
                : sourceParagraph.StartCharacter;
            int paragraphEndCharacter = pendingBookmarks != null && pendingBookmarks.Count > 0
                ? Math.Max(sourceParagraph.EndCharacter, pendingBookmarkEndCharacter)
                : sourceParagraph.EndCharacter;

            if (sourceParagraph.Runs.Count == 0) {
                WordParagraph emptyParagraph = cell.AddParagraph(removeExistingParagraphs: removeExistingParagraphs);
                emptyParagraph._paragraph.RemoveAllChildren<Run>();
                ApplyLegacyDocParagraphFormatting(emptyParagraph, sourceParagraph.Format, styleSheet);
                LegacyDocBookmarkProjection.Create(paragraphBookmarks, paragraphStartCharacter, paragraphEndCharacter).EmitRemaining(emptyParagraph._paragraph);
                return;
            }

            if (paragraphBookmarks.Count > 0) {
                WordParagraph bookmarkedParagraph = cell.AddParagraph(removeExistingParagraphs: removeExistingParagraphs);
                bookmarkedParagraph._paragraph.RemoveAllChildren<Run>();
                ApplyLegacyDocParagraphFormatting(bookmarkedParagraph, sourceParagraph.Format, styleSheet);
                LegacyDocBookmarkProjection bookmarks = LegacyDocBookmarkProjection.Create(paragraphBookmarks, paragraphStartCharacter, paragraphEndCharacter);
                AddLegacyDocRuns(bookmarkedParagraph, sourceParagraph.Runs, notes, bookmarks);
                bookmarks.EmitRemaining(bookmarkedParagraph._paragraph);
                return;
            }

            LegacyDocTextRun firstRun = sourceParagraph.Runs[0];
            WordParagraph paragraph;
            int remainingRunStartIndex;
            if (ContainsLegacyDocSpecialRunCharacter(firstRun.Text)
                || firstRun.HyperlinkTarget.HasValue
                || firstRun.FieldKind != LegacyDocFieldKind.None
                || firstRun.Picture != null) {
                paragraph = cell.AddParagraph(string.Empty, removeExistingParagraphs: removeExistingParagraphs);
                remainingRunStartIndex = 0;
            } else {
                paragraph = cell.AddParagraph(firstRun.Text, removeExistingParagraphs: removeExistingParagraphs);
                remainingRunStartIndex = 1;
            }

            ApplyLegacyDocParagraphFormatting(paragraph, sourceParagraph.Format, styleSheet);
            if (remainingRunStartIndex == 1) ApplyLegacyDocRunFormatting(paragraph, firstRun);
            AddLegacyDocRuns(paragraph, sourceParagraph.Runs, remainingRunStartIndex, notes);
        }

        private static IReadOnlyList<LegacyDocBookmark> MergeLegacyDocTableCellBookmarks(IReadOnlyList<LegacyDocBookmark> paragraphBookmarks, IReadOnlyList<LegacyDocBookmark>? pendingBookmarks) {
            if (pendingBookmarks == null || pendingBookmarks.Count == 0) {
                return paragraphBookmarks;
            }

            if (paragraphBookmarks.Count == 0) {
                return pendingBookmarks;
            }

            return pendingBookmarks
                .Concat(paragraphBookmarks)
                .Distinct()
                .ToArray();
        }

    }
}
