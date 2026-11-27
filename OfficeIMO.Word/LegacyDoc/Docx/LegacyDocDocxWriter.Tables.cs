using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using static OfficeIMO.Word.LegacyDoc.Docx.LegacyDocXml;

namespace OfficeIMO.Word.LegacyDoc.Docx {
    internal sealed partial class LegacyDocDocxWriter {
        private const int MaxProjectedTableCells = 100_000;

        /// <summary>
        /// Builds a table for <paramref name="tableBlock"/> in <paramref name="part"/>, detached, and returns it, or
        /// NoNode when the block has no cells. Nested tables found in cells are built through the same path.
        /// </summary>
        private int AddTable(LegacyDocDocxPart part, LegacyDocTableBlock tableBlock, LegacyDocNoteProjection notes, bool projectNestedTables, int? availableWidthTwips = null) {
            int rowCount = tableBlock.Rows.Count;
            int columnCount = rowCount == 0 ? 0 : tableBlock.Rows.Max(row => row.Cells.Count);
            if (rowCount == 0 || columnCount == 0) return OfficeXmlTree.NoNode;
            if ((long)rowCount * columnCount > MaxProjectedTableCells) {
                throw new InvalidDataException("The legacy DOC table exceeds the supported projected-cell limit.");
            }

            OfficeXmlTree tree = part.Tree;
            int defaultColumnWidth = DefaultColumnWidthTwips(availableWidthTwips ?? SectionTextWidthTwips(), columnCount);
            int table = Element(tree, "tbl");
            int properties = Append(tree, table, Element(tree, "tblPr"));
            Append(tree, properties, Element(tree, "tblStyle", "TableNormal"));
            int tableWidth = Append(tree, properties, Element(tree, "tblW"));
            Set(tree, tableWidth, "w", "0");
            Set(tree, tableWidth, "type", "auto");
            Append(tree, properties, CreateTableLook(tree));
            int grid = Append(tree, table, Element(tree, "tblGrid"));
            var cells = new int[rowCount][];
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                int row = Append(tree, table, Element(tree, "tr"));
                cells[rowIndex] = new int[columnCount];
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                    int cell = Append(tree, row, Element(tree, "tc"));
                    int cellProperties = Append(tree, cell, Element(tree, "tcPr"));
                    Append(tree, cellProperties, CreateWidth(tree, "tcW", defaultColumnWidth, "dxa"));
                    int paragraph = Append(tree, cell, Element(tree, "p"));
                    Append(tree, paragraph, Element(tree, "pPr"));
                    Append(tree, paragraph, Element(tree, "r"));
                    cells[rowIndex][columnIndex] = cell;
                }
            }

            ApplyTableBorderDefaults(tree, table, cells, tableBlock.Rows.Select(row =>
                row.TableBorders.WithDefaults(_styleSheet.ResolveTableBorders(row.TableStyleIndex))).ToArray());

            LegacyDocTableAlignment? tableAlignment = tableBlock.Rows.Select(row => row.TableAlignment).FirstOrDefault(alignment => alignment.HasValue);
            if (tableAlignment != null) {
                string? alignment = tableAlignment.Value switch {
                    LegacyDocTableAlignment.Left => "left",
                    LegacyDocTableAlignment.Center => "center",
                    LegacyDocTableAlignment.Right => "right",
                    _ => null
                };
                if (alignment != null) Replace(tree, properties, Element(tree, "jc", alignment), TablePropertiesOrder);
            }

            int? tableLeftIndentTwips = tableBlock.Rows.Select(row => row.TableLeftIndentTwips).FirstOrDefault(indent => indent.HasValue);
            if (tableLeftIndentTwips != null) Replace(tree, properties, CreateWidth(tree, "tblInd", tableLeftIndentTwips.Value, "dxa"), TablePropertiesOrder);

            LegacyDocTablePreferredWidth? tablePreferredWidth = tableBlock.Rows.Select(row => row.TablePreferredWidth).FirstOrDefault(width => width.HasValue);
            if (tablePreferredWidth != null) {
                switch (tablePreferredWidth.Value.Unit) {
                    case LegacyDocTablePreferredWidthUnit.Auto:
                        Replace(tree, properties, CreateWidth(tree, "tblW", 0, "auto"), TablePropertiesOrder);
                        break;
                    case LegacyDocTablePreferredWidthUnit.Percent:
                        Replace(tree, properties, CreateWidth(tree, "tblW", tablePreferredWidth.Value.Value, "pct"), TablePropertiesOrder);
                        break;
                    case LegacyDocTablePreferredWidthUnit.Dxa:
                        Replace(tree, properties, CreateWidth(tree, "tblW", tablePreferredWidth.Value.Value, "dxa"), TablePropertiesOrder);
                        break;
                }
            }

            bool? tableAutofit = tableBlock.Rows.Select(row => row.TableAutofit).FirstOrDefault(autofit => autofit.HasValue);
            // Binary DOC defaults to fixed columns when sprmTFAutofit is absent; the DOCX table must not acquire the
            // DOCX AutoFit default.
            int layout = Element(tree, "tblLayout");
            Set(tree, layout, "type", tableAutofit == true ? "autofit" : "fixed");
            Replace(tree, properties, layout, TablePropertiesOrder);

            int? tableCellSpacingTwips = tableBlock.Rows.Select(row => row.DefaultCellSpacingTwips).FirstOrDefault(spacing => spacing.HasValue);
            if (tableCellSpacingTwips != null) {
                Replace(tree, properties, CreateWidth(tree, "tblCellSpacing", checked((short)tableCellSpacingTwips.Value), "dxa"), TablePropertiesOrder);
            }

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                LegacyDocTableRow sourceRow = tableBlock.Rows[rowIndex];
                int row = tree.ParentOf(cells[rowIndex][0]);
                ApplyTableRowFormatting(tree, row, sourceRow);
                for (int columnIndex = 0; columnIndex < sourceRow.Cells.Count && columnIndex < columnCount; columnIndex++) {
                    int cell = cells[rowIndex][columnIndex];
                    int cellProperties = Child(tree, cell, "tcPr");
                    AddTableCell(part, cell, sourceRow.Cells[columnIndex], notes, projectNestedTables);
                    if (columnIndex < sourceRow.CellWidthsTwips.Count) {
                        Replace(tree, cellProperties, CreateWidth(tree, "tcW", sourceRow.CellWidthsTwips[columnIndex], "dxa"), TableCellPropertiesOrder);
                    }

                    if (columnIndex < sourceRow.CellHorizontalMerges.Count) {
                        string? merge = sourceRow.CellHorizontalMerges[columnIndex] switch {
                            LegacyDocTableCellHorizontalMerge.Restart => "restart",
                            LegacyDocTableCellHorizontalMerge.Continue => "continue",
                            _ => null
                        };
                        if (merge != null) Replace(tree, cellProperties, Element(tree, "hMerge", merge), TableCellPropertiesOrder);
                    }

                    if (columnIndex < sourceRow.CellVerticalMerges.Count) {
                        string? merge = sourceRow.CellVerticalMerges[columnIndex] switch {
                            LegacyDocTableCellVerticalMerge.Restart => "restart",
                            LegacyDocTableCellVerticalMerge.Continue => "continue",
                            _ => null
                        };
                        if (merge != null) Replace(tree, cellProperties, Element(tree, "vMerge", merge), TableCellPropertiesOrder);
                    }

                    if (columnIndex < sourceRow.CellVerticalAlignments.Count) {
                        string? alignment = sourceRow.CellVerticalAlignments[columnIndex] switch {
                            LegacyDocTableCellVerticalAlignment.Center => "center",
                            LegacyDocTableCellVerticalAlignment.Bottom => "bottom",
                            _ => null
                        };
                        if (alignment != null) Replace(tree, cellProperties, Element(tree, "vAlign", alignment), TableCellPropertiesOrder);
                    }

                    if (columnIndex < sourceRow.CellTextDirections.Count) {
                        string? direction = sourceRow.CellTextDirections[columnIndex] switch {
                            LegacyDocTableCellTextDirection.TopToBottomRightToLeft => "tbRl",
                            LegacyDocTableCellTextDirection.BottomToTopLeftToRight => "btLr",
                            LegacyDocTableCellTextDirection.LeftToRightTopToBottomRotated => "lrTbV",
                            LegacyDocTableCellTextDirection.TopToBottomRightToLeftRotated => "tbRlV",
                            _ => null
                        };
                        if (direction != null) Replace(tree, cellProperties, Element(tree, "textDirection", direction), TableCellPropertiesOrder);
                    }

                    if (columnIndex < sourceRow.CellFitTexts.Count && sourceRow.CellFitTexts[columnIndex]) {
                        Replace(tree, cellProperties, Element(tree, "tcFitText", "on"), TableCellPropertiesOrder);
                    }

                    if (columnIndex < sourceRow.CellNoWraps.Count && sourceRow.CellNoWraps[columnIndex]) {
                        Replace(tree, cellProperties, Element(tree, "noWrap"), TableCellPropertiesOrder);
                    }

                    if (columnIndex < sourceRow.CellHideMarks.Count && sourceRow.CellHideMarks[columnIndex]) {
                        Replace(tree, cellProperties, Element(tree, "hideMark", "on"), TableCellPropertiesOrder);
                    }

                    if (columnIndex < sourceRow.CellMargins.Count) ApplyTableCellMargins(tree, cellProperties, sourceRow.CellMargins[columnIndex]);
                    if (columnIndex < sourceRow.CellShadings.Count && !string.IsNullOrEmpty(sourceRow.CellShadings[columnIndex].FillColorHex)) {
                        int shading = Element(tree, "shd");
                        Set(tree, shading, "val", "clear");
                        Set(tree, shading, "fill", sourceRow.CellShadings[columnIndex].FillColorHex!);
                        Replace(tree, cellProperties, shading, TableCellPropertiesOrder);
                    }

                    if (columnIndex < sourceRow.CellBorders.Count) ApplyTableCellBorders(tree, cellProperties, sourceRow.CellBorders[columnIndex]);
                }
            }

            WriteTableGrid(tree, grid, cells, tableBlock, defaultColumnWidth);
            AddTableRowBoundaryBookmarks(tree, table, cells, tableBlock);
            return table;
        }

        private static int CreateTableLook(OfficeXmlTree tree) {
            return Element(tree, "tblLook", "04A0");
        }

        private static int CreateWidth(OfficeXmlTree tree, string name, int width, string type) {
            int element = Element(tree, name);
            Set(tree, element, "w", Text(width));
            Set(tree, element, "type", type);
            return element;
        }

        private const int DefaultTableTextWidthTwips = 9360;
        private const int DefaultCellHorizontalMarginsTwips = 216;
        private const int MinimumDefaultColumnWidthTwips = 360;

        /// <summary>The text width of the last section as it stands now, which new tables span by default.</summary>
        private int SectionTextWidthTwips() {
            OfficeXmlTree tree = _main.Tree;
            int sectionProperties = FinalSectionProperties();
            int pageSize = Child(tree, sectionProperties, "pgSz");
            if (pageSize == OfficeXmlTree.NoNode || !long.TryParse(Get(tree, pageSize, "w"), out long width)) return DefaultTableTextWidthTwips;
            int pageMargin = Child(tree, sectionProperties, "pgMar");
            long Margin(string name) => pageMargin != OfficeXmlTree.NoNode && long.TryParse(Get(tree, pageMargin, name), out long value) ? value : 0;
            width -= Margin("left") + Margin("right") + Margin("gutter");
            return width > MinimumDefaultColumnWidthTwips ? (int)Math.Min(width, int.MaxValue) : DefaultTableTextWidthTwips;
        }

        /// <summary>The text width of a cell a nested table is added to: its width less the default cell margins.</summary>
        private int CellTextWidthTwips(OfficeXmlTree tree, int cell) {
            int properties = Child(tree, cell, "tcPr");
            int width = properties == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : Child(tree, properties, "tcW");
            if (width != OfficeXmlTree.NoNode && Get(tree, width, "type") == "dxa" && int.TryParse(Get(tree, width, "w"), out int twips)
                && twips - DefaultCellHorizontalMarginsTwips > MinimumDefaultColumnWidthTwips) {
                return twips - DefaultCellHorizontalMarginsTwips;
            }

            return SectionTextWidthTwips();
        }

        private static int DefaultColumnWidthTwips(int textWidthTwips, int columns) =>
            Math.Max(MinimumDefaultColumnWidthTwips, textWidthTwips / Math.Max(1, columns));

        /// <summary>Writes one grid column per cell of the widest row, as wide as that row's cells.</summary>
        private static void WriteTableGrid(OfficeXmlTree tree, int grid, int[][] cells, LegacyDocTableBlock tableBlock, int defaultColumnWidth) {
            int columns = cells[0].Length;
            var widths = new int[columns];
            for (int column = 0; column < columns; column++) widths[column] = defaultColumnWidth;
            LegacyDocTableRow? widest = tableBlock.Rows.OrderByDescending(row => row.CellWidthsTwips.Count).FirstOrDefault();
            if (widest != null) {
                for (int column = 0; column < columns && column < widest.CellWidthsTwips.Count; column++) widths[column] = widest.CellWidthsTwips[column];
            }

            foreach (int width in widths) {
                int column = Element(tree, "gridCol");
                Set(tree, column, "w", Text(width));
                Append(tree, grid, column);
            }
        }

        private static void ApplyTableRowFormatting(OfficeXmlTree tree, int row, LegacyDocTableRow sourceRow) {
            if (sourceRow.RowHeightTwips != null) {
                int properties = RowProperties(tree, row);
                int height = Child(tree, properties, "trHeight");
                if (height == OfficeXmlTree.NoNode) {
                    height = Element(tree, "trHeight");
                    tree.InsertBefore(properties, height, tree.FirstChildOf(properties));
                }

                Set(tree, height, "val", Text(sourceRow.RowHeightTwips.Value));
                Set(tree, height, "hRule", sourceRow.RowHeightIsExact ? "exact" : "atLeast");
            }

            // The row setters put these first, ahead of the height.
            if (sourceRow.RowCantSplit == true && Child(tree, RowProperties(tree, row), "cantSplit") == OfficeXmlTree.NoNode) {
                int properties = RowProperties(tree, row);
                tree.InsertBefore(properties, Element(tree, "cantSplit"), tree.FirstChildOf(properties));
            }

            if (sourceRow.RowIsHeader == true && Child(tree, RowProperties(tree, row), "tblHeader") == OfficeXmlTree.NoNode) {
                int properties = RowProperties(tree, row);
                tree.InsertBefore(properties, Element(tree, "tblHeader", "on"), tree.FirstChildOf(properties));
            }
        }

        private static int RowProperties(OfficeXmlTree tree, int row) {
            int properties = Child(tree, row, "trPr");
            if (properties != OfficeXmlTree.NoNode) return properties;
            properties = Element(tree, "trPr");
            int first = tree.FirstChildOf(row);
            while (first != OfficeXmlTree.NoNode && Is(tree, first, "tblPrEx")) first = tree.NextSiblingOf(first);
            tree.InsertBefore(row, properties, first);
            return properties;
        }

        private static void ApplyTableCellMargins(OfficeXmlTree tree, int cellProperties, LegacyDocTableCellMargins margins) {
            if (margins.TopTwips == null && margins.RightTwips == null && margins.BottomTwips == null && margins.LeftTwips == null) return;
            int cellMargins = GetOrAdd(tree, cellProperties, "tcMar", TableCellPropertiesOrder);
            if (margins.TopTwips != null) Replace(tree, cellMargins, CreateWidth(tree, "top", checked((short)margins.TopTwips.Value), "dxa"), MarginsOrder);
            if (margins.RightTwips != null) Replace(tree, cellMargins, CreateWidth(tree, "right", checked((short)margins.RightTwips.Value), "dxa"), MarginsOrder);
            if (margins.BottomTwips != null) Replace(tree, cellMargins, CreateWidth(tree, "bottom", checked((short)margins.BottomTwips.Value), "dxa"), MarginsOrder);
            if (margins.LeftTwips != null) Replace(tree, cellMargins, CreateWidth(tree, "left", checked((short)margins.LeftTwips.Value), "dxa"), MarginsOrder);
        }

        private static void ApplyTableCellBorders(OfficeXmlTree tree, int cellProperties, LegacyDocTableCellBorders borders) {
            ApplyTableCellBorder(tree, cellProperties, "top", borders.Top);
            ApplyTableCellBorder(tree, cellProperties, "left", borders.Left);
            ApplyTableCellBorder(tree, cellProperties, "bottom", borders.Bottom);
            ApplyTableCellBorder(tree, cellProperties, "right", borders.Right);
        }

        private static void ApplyTableCellBorder(OfficeXmlTree tree, int cellProperties, string edge, LegacyDocTableCellBorder border) {
            string? style = TableCellBorderStyleToken(border.Style);
            if (style == null) return;
            int cellBorders = GetOrAdd(tree, cellProperties, "tcBorders", TableCellPropertiesOrder);
            // A specified edge replaces the earlier row default as a complete value.
            int element = Element(tree, edge, style);
            if (border.ColorHex != null) Set(tree, element, "color", border.ColorHex);
            if (border.SizeEighthPoints > 0) Set(tree, element, "sz", Text(border.SizeEighthPoints));
            if (border.SpacePoints > 0) Set(tree, element, "space", Text(border.SpacePoints));
            Replace(tree, cellBorders, element, BordersOrder);
        }

        private static string? TableCellBorderStyleToken(LegacyDocTableCellBorderStyle style) => style switch {
            LegacyDocTableCellBorderStyle.ExplicitNone => "nil",
            LegacyDocTableCellBorderStyle.Single => "single",
            LegacyDocTableCellBorderStyle.Double => "double",
            LegacyDocTableCellBorderStyle.Dotted => "dotted",
            LegacyDocTableCellBorderStyle.Dashed => "dashed",
            _ => null
        };

        private static void ApplyTableBorderDefaults(OfficeXmlTree tree, int table, int[][] cells, LegacyDocTableBorders[] rows) {
            if (rows.Length == 0) return;
            // A DOC operand belongs to a row. Only common edges can become table-wide defaults.
            LegacyDocTableCellBorder Common(Func<LegacyDocTableBorders, LegacyDocTableCellBorder> select) {
                LegacyDocTableCellBorder first = select(rows[0]);
                return rows.All(row => select(row).Equals(first)) ? first : default;
            }

            var source = new LegacyDocTableBorders(rows[0].Top, Common(row => row.Left),
                rows[rows.Length - 1].Bottom, Common(row => row.Right),
                Common(row => row.InsideHorizontal), Common(row => row.InsideVertical));
            int properties = Child(tree, table, "tblPr");
            int borders = Element(tree, "tblBorders");
            AppendTableBorder(tree, borders, "top", source.Top);
            AppendTableBorder(tree, borders, "left", source.Left);
            AppendTableBorder(tree, borders, "bottom", source.Bottom);
            AppendTableBorder(tree, borders, "right", source.Right);
            AppendTableBorder(tree, borders, "insideH", source.InsideHorizontal);
            AppendTableBorder(tree, borders, "insideV", source.InsideVertical);
            Replace(tree, properties, borders, TablePropertiesOrder);

            // A borderless row does not erase the visible edge supplied by its neighbor.
            LegacyDocTableCellBorder SharedHorizontal(LegacyDocTableCellBorder actual, LegacyDocTableCellBorder adjacent) =>
                (!actual.HasAny || actual.Style == LegacyDocTableCellBorderStyle.ExplicitNone)
                    && adjacent.HasAny && adjacent.Style != LegacyDocTableCellBorderStyle.ExplicitNone
                    ? adjacent : actual;
            for (int row = 0; row < rows.Length && row < cells.Length; row++) {
                int[] rowCells = cells[row];
                for (int column = 0; column < rowCells.Length; column++) {
                    LegacyDocTableCellBorder Override(LegacyDocTableCellBorder actual, LegacyDocTableCellBorder common) {
                        if (actual.Equals(common)) return default;
                        return actual.HasAny ? actual : new LegacyDocTableCellBorder(LegacyDocTableCellBorderStyle.ExplicitNone, null, 0, 0);
                    }

                    ApplyTableCellBorders(tree, Child(tree, rowCells[column], "tcPr"), new LegacyDocTableCellBorders(
                        Override(row == 0 ? rows[row].Top : SharedHorizontal(rows[row].InsideHorizontal, rows[row - 1].InsideHorizontal),
                            row == 0 ? source.Top : source.InsideHorizontal),
                        Override(column == 0 ? rows[row].Left : rows[row].InsideVertical,
                            column == 0 ? source.Left : source.InsideVertical),
                        Override(row == rows.Length - 1 ? rows[row].Bottom : SharedHorizontal(rows[row].InsideHorizontal, rows[row + 1].InsideHorizontal),
                            row == rows.Length - 1 ? source.Bottom : source.InsideHorizontal),
                        Override(column == rowCells.Length - 1 ? rows[row].Right : rows[row].InsideVertical,
                            column == rowCells.Length - 1 ? source.Right : source.InsideVertical)));
                }
            }
        }

        private static void AppendTableBorder(OfficeXmlTree tree, int borders, string edge, LegacyDocTableCellBorder source) {
            string? style = TableCellBorderStyleToken(source.Style);
            if (style == null) return;
            int border = Element(tree, edge, style);
            if (style != "nil") Set(tree, border, "color", source.ColorHex ?? "auto");
            if (source.SizeEighthPoints > 0) Set(tree, border, "sz", Text(source.SizeEighthPoints));
            if (source.SpacePoints > 0) Set(tree, border, "space", Text(source.SpacePoints));
            Append(tree, borders, border);
        }

        // ---- Cells ----------------------------------------------------------------------------------------------

        private void AddTableCell(LegacyDocDocxPart part, int cell, LegacyDocTableCell sourceCell, LegacyDocNoteProjection notes, bool projectNestedTables) {
            var pendingBookmarks = new List<LegacyDocBookmark>();
            int pendingStart = int.MaxValue;
            int pendingEnd = int.MinValue;
            bool emittedParagraph = false;
            for (int index = 0; index < sourceCell.Paragraphs.Count; index++) {
                LegacyDocTableCellParagraph sourceParagraph = sourceCell.Paragraphs[index];
                if (string.IsNullOrEmpty(sourceParagraph.Text) && sourceParagraph.Bookmarks.Count > 0 && index + 1 < sourceCell.Paragraphs.Count) {
                    foreach (LegacyDocBookmark bookmark in sourceParagraph.Bookmarks) {
                        pendingBookmarks.Add(bookmark);
                        pendingStart = Math.Min(pendingStart, bookmark.StartCharacter);
                        pendingEnd = Math.Max(pendingEnd, bookmark.EndCharacter);
                    }

                    continue;
                }

                if (projectNestedTables && IsNestedTableParagraph(sourceParagraph)) {
                    index = AddNestedTable(part, cell, sourceCell.Paragraphs, index, notes, pendingBookmarks);
                    emittedParagraph = true;
                    pendingBookmarks.Clear();
                    pendingStart = int.MaxValue;
                    pendingEnd = int.MinValue;
                    continue;
                }

                AddTableCellParagraph(part, cell, sourceParagraph, notes, removeExistingParagraphs: !emittedParagraph, pendingBookmarks, pendingStart, pendingEnd);
                emittedParagraph = true;
                pendingBookmarks.Clear();
                pendingStart = int.MaxValue;
                pendingEnd = int.MinValue;
            }
        }

        private static void RemoveCellParagraphs(OfficeXmlTree tree, int cell) {
            foreach (int paragraph in Elements(tree, cell, "p")) tree.Remove(paragraph);
        }

        private void AddTableCellParagraph(LegacyDocDocxPart part, int cell, LegacyDocTableCellParagraph sourceParagraph, LegacyDocNoteProjection notes,
            bool removeExistingParagraphs, IReadOnlyList<LegacyDocBookmark> pendingBookmarks, int pendingStart, int pendingEnd) {
            OfficeXmlTree tree = part.Tree;
            IReadOnlyList<LegacyDocBookmark> paragraphBookmarks = pendingBookmarks.Count == 0
                ? sourceParagraph.Bookmarks
                : sourceParagraph.Bookmarks.Count == 0 ? pendingBookmarks.ToArray() : pendingBookmarks.Concat(sourceParagraph.Bookmarks).Distinct().ToArray();
            int startCharacter = pendingBookmarks.Count > 0 ? Math.Min(sourceParagraph.StartCharacter, pendingStart) : sourceParagraph.StartCharacter;
            int endCharacter = pendingBookmarks.Count > 0 ? Math.Max(sourceParagraph.EndCharacter, pendingEnd) : sourceParagraph.EndCharacter;
            if (removeExistingParagraphs) RemoveCellParagraphs(tree, cell);
            int paragraph = Append(tree, cell, Element(tree, "p"));
            Append(tree, paragraph, Element(tree, "pPr"));
            ApplyParagraphFormatting(part, paragraph, sourceParagraph.Format);
            var target = new RunTarget(part, paragraph, paragraph);
            if (sourceParagraph.Runs.Count == 0) {
                LegacyDocBookmarkProjection.Create(paragraphBookmarks, startCharacter, endCharacter).EmitRemaining(tree, paragraph);
                return;
            }

            if (paragraphBookmarks.Count > 0) {
                var bookmarks = LegacyDocBookmarkProjection.Create(paragraphBookmarks, startCharacter, endCharacter);
                AddRuns(target, sourceParagraph.Runs, 0, notes, bookmarks);
                bookmarks.EmitRemaining(tree, paragraph);
                return;
            }

            // The cell API creates the paragraph with a text run: the first source run when it is plain text, or an
            // empty one. That run stays the paragraph's current run: it takes the first run's formatting (unless it is
            // empty) and any picture added later in the paragraph.
            LegacyDocTextRun firstRun = sourceParagraph.Runs[0];
            bool plainFirstRun = !(ContainsSpecialRunCharacter(firstRun.Text) || firstRun.HyperlinkTarget.HasValue
                || firstRun.FieldKind != LegacyDocFieldKind.None || firstRun.Picture != null);
            int currentRun = Append(tree, paragraph, CreateRun(tree, TextElement(tree, plainFirstRun ? firstRun.Text : string.Empty)));
            if (plainFirstRun && firstRun.Text.Length > 0) ApplyRunFormatting(target, currentRun, firstRun);
            AddRuns(new RunTarget(part, paragraph, paragraph, currentRun), sourceParagraph.Runs, plainFirstRun ? 1 : 0, notes, LegacyDocBookmarkProjection.Empty);
        }

        private void AddTableRowBoundaryBookmarks(OfficeXmlTree tree, int table, int[][] cells, LegacyDocTableBlock tableBlock) {
            int rowCount = Math.Min(cells.Length, tableBlock.Rows.Count);
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                LegacyDocTableRow sourceRow = tableBlock.Rows[rowIndex];
                if (sourceRow.BookmarksBefore.Count == 0) continue;
                int row = tree.ParentOf(cells[rowIndex][0]);
                foreach (LegacyDocBookmark bookmark in sourceRow.BookmarksBefore.OrderBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                    tree.InsertBefore(table, LegacyDocBookmarkProjection.CreateBookmarkStart(tree, bookmark.ProjectionId, bookmark.Name), row);
                    tree.InsertBefore(table, LegacyDocBookmarkProjection.CreateBookmarkEnd(tree, bookmark.ProjectionId), row);
                }
            }
        }

        /// <summary>Places the bookmarks that start or end exactly at a table around it.</summary>
        private static void AddTableBlockBookmarks(OfficeXmlTree tree, int table, LegacyDocTableBlock tableBlock) {
            int parent = tree.ParentOf(table);
            if (tableBlock.Bookmarks.Count == 0 || parent == OfficeXmlTree.NoNode) return;
            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => bookmark.IsZeroLength && bookmark.StartCharacter == tableBlock.StartCharacter)
                .OrderBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                tree.InsertBefore(parent, LegacyDocBookmarkProjection.CreateBookmarkStart(tree, bookmark.ProjectionId, bookmark.Name), table);
                tree.InsertBefore(parent, LegacyDocBookmarkProjection.CreateBookmarkEnd(tree, bookmark.ProjectionId), table);
            }

            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => !bookmark.IsZeroLength && bookmark.StartCharacter == tableBlock.StartCharacter)
                .OrderByDescending(bookmark => bookmark.EndCharacter)
                .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                tree.InsertBefore(parent, LegacyDocBookmarkProjection.CreateBookmarkStart(tree, bookmark.ProjectionId, bookmark.Name), table);
            }

            int afterAnchor = table;
            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => bookmark.EndCharacter == tableBlock.EndCharacter && !bookmark.IsZeroLength)
                .OrderByDescending(bookmark => bookmark.StartCharacter)
                .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                afterAnchor = InsertAfter(tree, parent, LegacyDocBookmarkProjection.CreateBookmarkEnd(tree, bookmark.ProjectionId), afterAnchor);
            }

            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => bookmark.IsZeroLength && bookmark.StartCharacter == tableBlock.EndCharacter)
                .OrderBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                afterAnchor = InsertAfter(tree, parent, LegacyDocBookmarkProjection.CreateBookmarkStart(tree, bookmark.ProjectionId, bookmark.Name), afterAnchor);
                afterAnchor = InsertAfter(tree, parent, LegacyDocBookmarkProjection.CreateBookmarkEnd(tree, bookmark.ProjectionId), afterAnchor);
            }
        }

        private static int InsertAfter(OfficeXmlTree tree, int parent, int node, int anchor) {
            tree.InsertBefore(parent, node, tree.NextSiblingOf(anchor));
            return node;
        }

        // ---- Nested tables --------------------------------------------------------------------------------------

        private static bool IsNestedTableParagraph(LegacyDocTableCellParagraph paragraph) =>
            paragraph.Format.MaximumTableDepth > 1
            || paragraph.Format.HasInnerTableCellMarker
            || paragraph.Format.HasInnerTableTerminatingParagraphMarker;

        private int AddNestedTable(LegacyDocDocxPart part, int hostCell, IReadOnlyList<LegacyDocTableCellParagraph> paragraphs, int startIndex,
            LegacyDocNoteProjection notes, IReadOnlyList<LegacyDocBookmark> pendingBookmarks) {
            var rows = new List<List<List<LegacyDocTableCellParagraph>>>();
            var rowFormats = new List<LegacyDocParagraphFormat>();
            var currentRow = new List<List<LegacyDocTableCellParagraph>>();
            var currentCell = new List<LegacyDocTableCellParagraph>();
            LegacyDocParagraphFormat lastRowFormat = LegacyDocParagraphFormat.Default;
            int index = startIndex;
            for (; index < paragraphs.Count; index++) {
                LegacyDocTableCellParagraph paragraph = paragraphs[index];
                if (!IsNestedTableParagraph(paragraph)) break;
                lastRowFormat = paragraph.Format;
                if (!paragraph.Format.HasInnerTableTerminatingParagraphMarker || paragraph.Runs.Count > 0 || paragraph.Bookmarks.Count > 0) {
                    currentCell.Add(paragraph);
                }

                if (paragraph.Format.HasInnerTableCellMarker) CloseNestedCell();
                if (paragraph.Format.HasInnerTableTerminatingParagraphMarker) {
                    if (currentCell.Count > 0 || currentRow.Count == 0) CloseNestedCell();
                    if (currentRow.Count > 0) {
                        rows.Add(currentRow);
                        rowFormats.Add(paragraph.Format);
                        currentRow = new List<List<LegacyDocTableCellParagraph>>();
                    }
                }
            }

            if (currentCell.Count > 0) CloseNestedCell();
            if (currentRow.Count > 0) {
                rows.Add(currentRow);
                rowFormats.Add(lastRowFormat);
            }

            if (rows.Count == 0) return startIndex;

            // Use the ordinary table writer so nested widths, row constraints, cell settings and borders follow the
            // same native DOC contract.
            var sourceRows = new List<LegacyDocTableRow>();
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                LegacyDocParagraphFormat format = rowFormats[rowIndex];
                LegacyDocTableCell[] cells = rows[rowIndex].Select(cell => new LegacyDocTableCell(cell)).ToArray();
                sourceRows.Add(new LegacyDocTableRow(
                    cells, format.TableCellWidthsTwips, format.TableLeftIndentTwips,
                    format.TableRowHeightTwips, format.TableRowHeightIsExact,
                    format.TableRowCantSplit, format.TableRowIsHeader, format.TableAlignment,
                    format.TableCellHorizontalMerges, format.TableCellVerticalMerges,
                    format.TableCellVerticalAlignments, format.TableCellTextDirections,
                    format.TableCellFitTexts, format.TableCellNoWraps, format.TableCellHideMarks,
                    format.GetTableCellMarginsForCellCount(cells.Length),
                    format.GetTableCellShadingsForCellCount(cells.Length),
                    format.GetTableCellBordersForCellCount(cells.Length),
                    format.DefaultTableCellSpacingTwips, format.TablePreferredWidth, format.TableAutofit,
                    tableStyleIndex: format.TableStyleIndex, tableBorders: format.TableBorders));
            }

            OfficeXmlTree tree = part.Tree;
            int nestedTable = AddTable(part, new LegacyDocTableBlock(sourceRows, 0, 0), notes, projectNestedTables: false, availableWidthTwips: CellTextWidthTwips(tree, hostCell));
            if (nestedTable != OfficeXmlTree.NoNode) {
                // As the cell's table API does: drop a preceding empty paragraph, add the table, and end the cell with
                // the empty zero-spaced paragraph Word requires after a table.
                List<int> existing = Elements(tree, hostCell, "p");
                if (existing.Count > 0 && !HasMeaningfulParagraphContent(tree, existing[existing.Count - 1])) tree.Remove(existing[existing.Count - 1]);
                Append(tree, hostCell, nestedTable);
                int trailing = Append(tree, hostCell, Element(tree, "p"));
                int spacing = Append(tree, Append(tree, trailing, Element(tree, "pPr")), Element(tree, "spacing"));
                Set(tree, spacing, "before", "0");
                Set(tree, spacing, "after", "0");
                Set(tree, spacing, "line", "0");
                AddPendingBookmarksAroundNestedTable(tree, hostCell, nestedTable, pendingBookmarks);
            }

            return index - 1;

            void CloseNestedCell() {
                currentRow.Add(currentCell);
                currentCell = new List<LegacyDocTableCellParagraph>();
            }
        }

        private static bool HasMeaningfulParagraphContent(OfficeXmlTree tree, int paragraph) {
            int properties = Child(tree, paragraph, "pPr");
            if (properties != OfficeXmlTree.NoNode && (HasElements(tree, properties) || tree.AttributeCountOf(properties) > 0)) return true;
            foreach (int element in Descendants(tree, paragraph)) {
                if (Is(tree, element, "t")) {
                    if (!string.IsNullOrWhiteSpace(tree.GetInnerText(element))) return true;
                    continue;
                }

                if (Is(tree, element, "pPr") || Is(tree, element, "r") || Is(tree, element, "rPr")) continue;
                // Drawings, fields, bookmarks, hyperlinks and other non-text content are data.
                return true;
            }

            return false;
        }

        private static void AddPendingBookmarksAroundNestedTable(OfficeXmlTree tree, int parent, int nestedTable, IReadOnlyList<LegacyDocBookmark> pendingBookmarks) {
            if (pendingBookmarks.Count == 0) return;
            foreach (LegacyDocBookmark bookmark in pendingBookmarks
                .OrderByDescending(bookmark => bookmark.EndCharacter)
                .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                tree.InsertBefore(parent, LegacyDocBookmarkProjection.CreateBookmarkStart(tree, bookmark.ProjectionId, bookmark.Name), nestedTable);
            }

            int afterAnchor = nestedTable;
            foreach (LegacyDocBookmark bookmark in pendingBookmarks
                .OrderBy(bookmark => bookmark.StartCharacter)
                .ThenByDescending(bookmark => bookmark.EndCharacter)
                .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                afterAnchor = InsertAfter(tree, parent, LegacyDocBookmarkProjection.CreateBookmarkEnd(tree, bookmark.ProjectionId), afterAnchor);
            }
        }
    }
}
