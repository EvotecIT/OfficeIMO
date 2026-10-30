using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        /// <summary>Projects source rows, styles and content once for both ordinary and contained table flow.</summary>
        private static PdfCore.TableBlock? BuildNativeTableBlock(WordTable table, Func<WordParagraph, (int Level, string Marker)?>? getMarker, NativeNoteNumbering? footnoteNumbersById, WordPdfExportOptions? options, double? contentWidth, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap, bool supportsPositionedTables, int tableNestingDepth = 0) {
            EnsureNativeTableDepth(tableNestingDepth);
            TableLayout layout = TableLayoutCache.GetLayout(table);
            bool hasExplicitDefaultTableStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true;
            NativeTableStyleDefaults tableStyleDefaults = GetNativeTableStyleDefaults(
                table,
                nativeDefaults,
                ignoreFallbackTableStyle: hasExplicitDefaultTableStyle);
            bool hasNestedTables = table._table.Descendants<W.Table>().Any();
            bool preservesNestedTables = hasNestedTables &&
                nativeDefaults.RunningContentContext == null &&
                (tableNestingDepth > 0 || supportsPositionedTables || nativeDefaults.NoteContentSection != null) &&
                CanProjectNativeContainedTable(table, options, nativeDefaults, nativeFontMap, tableNestingDepth);
            if (hasNestedTables && !preservesNestedTables && options != null)
                AddNativeNestedTableLayoutWarning(options, "table");
            bool usesAutoFitLayout = ShouldUseNativeAutoFitTableLayout(
                table._table.GetFirstChild<W.TableProperties>(), tableStyleDefaults);
            var rows = new List<PdfCore.PdfTableCell[]>();
            var cellFills = new Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor>();
            var directCellBorders = new Dictionary<(int Row, int Column), WordTableCellBorder>();
            var cellPaddings = new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>();
            var cellAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>();
            var cellVerticalAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>();
            NativeTableColumnAlignments tableAlignments = CreateNativeTableColumnAlignments(layout);
            List<PdfCore.PdfColumnAlign>? horizontalAlignments = tableAlignments.Horizontal;
            List<PdfCore.PdfCellVerticalAlign>? verticalAlignments = tableAlignments.Vertical;
            int tableColumnCount = GetNativeTableColumnCount(layout);
            int repeatedHeaderRowCount = GetNativeTableRepeatedHeaderRowCount(table, layout.Rows.Count);
            int visualHeaderRowCount = GetNativeTableVisualHeaderRowCount(table, layout.Rows.Count, repeatedHeaderRowCount);
            int footerStartRowIndex = table.ConditionalFormattingLastRow == true && layout.Rows.Count > visualHeaderRowCount
                ? layout.Rows.Count - 1
                : layout.Rows.Count;
            for (int rowIndex = 0; rowIndex < layout.Rows.Count; rowIndex++) {
                IReadOnlyList<WordTableCell> row = layout.Rows[rowIndex];
                var nativeCells = new List<PdfCore.PdfTableCell>();
                int logicalColumnIndex = GetNativeTableRowStartColumn(layout, rowIndex);
                AddNativeTableGridBeforePlaceholders(nativeCells, logicalColumnIndex);
                for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                    WordTableCell cell = row[columnIndex];
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    // Vertical continuations still contribute margins to the physical row box,
                    // even though their text/content is rendered by the spanning anchor.
                    PdfCore.PdfCellPadding? padding = CreateNativeTableCellPadding(cell);
                    if (padding != null) {
                        cellPaddings[(rowIndex, logicalColumnIndex)] = padding;
                    }
                    if (IsNativeVerticalMergeContinuation(cell)) {
                        logicalColumnIndex += columnSpan;
                        continue;
                    }

                    NativeTableStyleDefaults cellStyleDefaults = GetNativeTableCellStyleDefaults(
                        table,
                        tableStyleDefaults,
                        rowIndex,
                        logicalColumnIndex,
                        columnSpan,
                        tableColumnCount,
                        visualHeaderRowCount,
                        footerStartRowIndex);
                    NativeTableCellEmbeddedContent embeddedContent = CreateNativeTableCellEmbeddedContent(cell, options,
                        interactiveControls: nativeDefaults.RunningContentContext == null);
                    NativeCellText cellText = CreateNativeCellText(
                        cell,
                        footnoteNumbersById,
                        nativeDefaults,
                        cellStyleDefaults,
                        nativeFontMap,
                        getMarker,
                        ignoreFallbackTableStyle: hasExplicitDefaultTableStyle,
                        options: options,
                        inlineImages: embeddedContent.InlineImages,
                        tableNestingDepth: tableNestingDepth,
                        preserveNestedBlocks: preservesNestedTables && cell._tableCell.Elements<W.Table>().Any());
                    int rowSpan = GetNativeCellRowSpan(cell);
                    nativeCells.Add(new PdfCore.PdfTableCell(
                        cellText.Runs,
                        cellText.Paragraphs,
                        columnSpan,
                        rowSpan: rowSpan,
                        checkBoxes: embeddedContent.CheckBoxes.Count == 0 ? null : embeddedContent.CheckBoxes,
                        formFields: embeddedContent.FormFields.Count == 0 ? null : embeddedContent.FormFields,
                        images: embeddedContent.Images.Count == 0 ? null : embeddedContent.Images,
                        // Word no-wrap changes automatic sizing. Fixed-layout tables and
                        // cells with absolute preferred widths still wrap their content.
                        noWrap: usesAutoFitLayout && WordTableWidthNative.CanDisableCellWrapping(cell) && !cell.WrapText,
                        textRotation: GetNativeCellTextRotation(cell.TextDirection),
                        blocks: cellText.Blocks) {
                        OrientedRowTextHeight = GetNativeOrientedCellMarkHeight(cell, cellText.Runs, embeddedContent,
                            nativeDefaults, cellStyleDefaults, nativeFontMap)
                    });

                    global::OfficeIMO.OfficeColor? fill =
                        ParseNativeColor(cell.ShadingFillColorHex) ??
                        cellStyleDefaults.CellFill;
                    if (fill.HasValue) {
                        cellFills[(rowIndex, logicalColumnIndex)] = fill.Value;
                    }

                    if (HasNativeDirectCellBorder(cell.Borders)) {
                        directCellBorders[(rowIndex, logicalColumnIndex)] = cell.Borders;
                    }

                    PdfCore.PdfColumnAlign cellAlignment = GetNativeCellHorizontalAlignment(cell);
                    if (cellAlignment != PdfCore.PdfColumnAlign.Left) {
                        cellAlignments[(rowIndex, logicalColumnIndex)] = cellAlignment;
                    }

                    PdfCore.PdfCellVerticalAlign? cellVerticalAlignment = ResolveNativeTableCellVerticalAlignment(cell, cellStyleDefaults);
                    if (cellVerticalAlignment.HasValue) {
                        cellVerticalAlignments[(rowIndex, logicalColumnIndex)] = cellVerticalAlignment.Value;
                    }

                    logicalColumnIndex += columnSpan;
                }

                AddNativeTableGridAfterPlaceholders(nativeCells, GetNativeTableRowTrailingColumnCount(layout, rowIndex));
                rows.Add(nativeCells.ToArray());
            }

            if (rows.Count == 0) {
                return null;
            }

            PdfCore.PdfTableStyle style = CreateNativeTableStyle(
                table,
                rows.Count,
                options,
                contentWidth,
                nativeDefaults,
                tableStyleDefaults,
                layout,
                nativeFontMap);
            if (table._tableProperties?.TablePositionProperties is { } tablePosition) {
                style.ConsumesVerticalFlow = false;
                if (!supportsPositionedTables && options != null) {
                    AddNativeExportWarning(options,
                        "native-positioned-table-wrap-approximation",
                        "table",
                        "Positioned table wrapping in a flow container without positioned-table support is approximated.");
                }
                if (supportsPositionedTables) {
                    style.Position = CreateNativeTablePosition(tablePosition,
                        table._tableProperties?.GetFirstChild<W.TableOverlap>()?.Val?.Value != W.TableOverlapValues.Never);
                    // tblpX is a placement coordinate, not a width reservation.
                    style.LeftIndent = 0;
                }
            }
            if (cellFills.Count > 0) {
                if (style.CellFills == null) {
                    style.CellFills = cellFills;
                } else {
                    foreach (var cellFill in cellFills) {
                        style.CellFills[cellFill.Key] = cellFill.Value;
                    }
                }
            }

            ApplyNativeDirectCellBorders(style, layout, directCellBorders);
            if (style.CellBorders != null && !UsesNativeSeparateTableCellBorders(table, style, tableStyleDefaults)) {
                ReconcileNativeHiddenSharedBorders(table, layout, tableStyleDefaults, style.HeaderRowCount, style.CellBorders, directCellBorders);
            }

            if (cellPaddings.Count > 0) {
                if (style.CellPaddings == null) {
                    style.CellPaddings = cellPaddings;
                } else {
                    var mergedPaddings = new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>(style.CellPaddings);
                    foreach (var cellPadding in cellPaddings) {
                        mergedPaddings[cellPadding.Key] = MergeNativeCellPadding(
                            mergedPaddings.TryGetValue(cellPadding.Key, out PdfCore.PdfCellPadding? existing) ? existing : null,
                            cellPadding.Value)!;
                    }

                    style.CellPaddings = mergedPaddings;
                }
            }

            if (cellAlignments.Count > 0) {
                style.CellAlignments = cellAlignments;
            }

            if (cellVerticalAlignments.Count > 0) {
                if (style.CellVerticalAlignments == null) {
                    style.CellVerticalAlignments = cellVerticalAlignments;
                } else {
                    var mergedVerticalAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>(style.CellVerticalAlignments);
                    foreach (var cellVerticalAlignment in cellVerticalAlignments) {
                        mergedVerticalAlignments[cellVerticalAlignment.Key] = cellVerticalAlignment.Value;
                    }

                    style.CellVerticalAlignments = mergedVerticalAlignments;
                }
            }

            ApplyNativeTablePercentageWidth(table, layout, style, contentWidth, tableStyleDefaults);
            ApplyNativeTableBorderFrame(table, layout, style, tableStyleDefaults);
            ApplyNativeUnspacedTableRowMargins(table, layout, style);
            // An explicit PDF table style retains its presentation placement.
            // Source compatibility geometry applies to the Word-owned style path.
            if (!hasExplicitDefaultTableStyle || ResolveNativeWordTableStyle(table, hasExplicitDefaultTableStyle) != null)
                ApplyNativeInlineTablePlacement(table, layout, style, tableStyleDefaults, containedTable: tableNestingDepth > 0);
            if (tableNestingDepth > 0) {
                // The containing cell is already a physical frame. Retain its
                // authored cell margins while compressing only grid flexibility.
                if (!style.AutoFitColumns) style.CompressFixedColumnsAboveMinimums = true;
                ApplyNativeContainedTablePercentageWidth(table, layout, style, tableStyleDefaults);
            }
            ApplyNativeFixedTableMinimumWidths(layout, style,
                includeConstrainedAutoFit: style.PreserveAuthoredPreferencesAtWidthLimit);
            ApplyNativeZeroSpacingCellMargins(layout, style);
            ApplyNativeFixedPercentageColumnPreferences(layout, style);
            ApplyNativeAutomaticTableFrameWidth(layout, style, contentWidth);
            ApplyNativeColumnWidths(table, layout, style, contentWidth);
            style = CreateNativeNoteTableStyle(table, style, nativeDefaults);

            if (horizontalAlignments != null) {
                style.Alignments = horizontalAlignments;
            }

            if (verticalAlignments != null) {
                style.VerticalAlignments = verticalAlignments;
            }

            return new PdfCore.TableBlock(rows, MapNativeTableAlignment(ResolveNativeTableAlignment(table, tableStyleDefaults)), style);
        }

    }
}
