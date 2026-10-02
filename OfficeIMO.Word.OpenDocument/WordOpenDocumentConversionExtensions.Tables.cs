using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    private static void ConvertTable(WordTableSnapshot source, OdtDocument targetDocument,
        WordOpenDocumentConversionOptions options, OdfImageValidationBudget imageValidationBudget,
        ref int hyperlinks, ref int images, ref int unsupportedImages,
        ref int bookmarks, NoteMappingStats notes) {
        int rows = Math.Max(1, source.RowCount);
        int columns = Math.Max(1, source.ColumnCount);
        OdtTable target = targetDocument.AddTable(rows, columns, source.Title);
        var covered = new bool[rows, columns];
        foreach (WordTableRowSnapshot row in source.Rows) {
            foreach (WordTableCellSnapshot cell in row.Cells) {
                int column = cell.ColumnIndex;
                if (row.RowIndex < 0 || row.RowIndex >= rows || column < 0 || column >= columns || covered[row.RowIndex, column]) continue;
                OdtTableCell targetCell = target.Cell(row.RowIndex, column);
                for (int paragraphIndex = 0; paragraphIndex < cell.Paragraphs.Count; paragraphIndex++) {
                    OdtParagraph targetParagraph = paragraphIndex == 0
                        ? targetCell.Paragraphs[0]
                        : targetCell.AddParagraph();
                    CopyParagraph(cell.Paragraphs[paragraphIndex], targetParagraph, options, imageValidationBudget, ref hyperlinks, ref images,
                        ref unsupportedImages, ref bookmarks, notes);
                }
                int rowSpan = Math.Min(cell.RowSpan, rows - row.RowIndex);
                int columnSpan = Math.Min(cell.ColumnSpan, columns - column);
                if (rowSpan > 1 || columnSpan > 1) {
                    target.Merge(row.RowIndex, column, rowSpan, columnSpan);
                    for (int y = 0; y < rowSpan; y++) for (int x = 0; x < columnSpan; x++)
                            if (x != 0 || y != 0) covered[row.RowIndex + y, column + x] = true;
                }
            }
        }
    }

    private static void ConvertTable(OdtTable source, WordDocument targetDocument,
        WordOpenDocumentConversionOptions options, CultureInfo textCaseCulture,
        ref int hyperlinks, ref int externalHyperlinks, ref int images,
        ref int bookmarks, ref int approximatedRuns, ref int approximatedBookmarkRanges, ref int unsupportedMeasurements,
        ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies,
        ref int mappedFields, ref int unsupportedFields,
        HashSet<System.Xml.Linq.XElement> handledUnsupportedFieldElements, NoteMappingStats notes,
        OdfConversionReport report, WordTableCell? parentCell = null) {
        int rows = Math.Max(1, source.Rows.Count);
        int columns = Math.Max(1, source.Rows.Select(row => row.Cells.Count).DefaultIfEmpty(1).Max());
        WordTable target = parentCell == null ? targetDocument.AddTable(rows, columns) : parentCell.AddTable(rows, columns);
        var merges = new List<(int Row, int Column, int RowSpan, int ColumnSpan)>();
        for (int row = 0; row < source.Rows.Count; row++) {
            IReadOnlyList<OdtTableCell> cells = source.Rows[row].Cells;
            for (int column = 0; column < cells.Count && column < columns; column++) {
                OdtTableCell cell = cells[column];
                if (cell.IsCovered) continue;
                WordTableCell targetCell = target.Rows[row].Cells[column];
                bool firstBlock = true;
                WordList? currentList = null;
                bool? currentOrdered = null;
                foreach (OdtContentBlock block in cell.ReadContentBlocks()) {
                    if (block.Table != null) {
                        ConvertTable(block.Table, targetDocument, options, textCaseCulture, ref hyperlinks, ref externalHyperlinks,
                            ref images, ref bookmarks, ref approximatedRuns, ref approximatedBookmarkRanges,
                            ref unsupportedMeasurements, ref approximatedFontFamilyLists, ref unsupportedFontFamilies,
                            ref mappedFields, ref unsupportedFields, handledUnsupportedFieldElements, notes, report, targetCell);
                        report.Add("nested-tables", OdfConversionMappingStatus.Converted, 1);
                        currentList = null;
                        currentOrdered = null;
                    } else {
                        WordParagraph targetParagraph;
                        if (block.IsListItem) {
                            bool ordered = block.IsOrderedList == true;
                            if (currentList == null || currentOrdered != ordered) {
                                currentList = targetCell.AddList(ordered ? WordListStyle.Numbered : WordListStyle.Bulleted);
                                currentOrdered = ordered;
                                report.Add("table-cell-lists", OdfConversionMappingStatus.Converted, 1);
                            }
                            targetParagraph = currentList.AddItem(null, Math.Max(0, Math.Min(8, block.ListLevel)));
                            if (block.ListLevel > 8) report.Add("list-levels", OdfConversionMappingStatus.Approximated, 1);
                        } else {
                            currentList = null;
                            currentOrdered = null;
                            targetParagraph = targetCell.AddParagraph(removeExistingParagraphs: firstBlock);
                            if (block.Paragraph!.IsHeading) targetParagraph.Style = HeadingStyle(block.Paragraph.HeadingLevel ?? 1);
                        }
                        CopyParagraph(block.Paragraph!, targetParagraph, options, textCaseCulture, ref hyperlinks,
                            ref externalHyperlinks, ref images, ref bookmarks, ref approximatedRuns,
                            ref approximatedBookmarkRanges, ref unsupportedMeasurements,
                            ref approximatedFontFamilyLists, ref unsupportedFontFamilies,
                            ref mappedFields, ref unsupportedFields, handledUnsupportedFieldElements, notes);
                    }
                    firstBlock = false;
                }
                if (cell.RowSpan > 1 || cell.ColumnSpan > 1) merges.Add((row, column, cell.RowSpan, cell.ColumnSpan));
            }
        }
        foreach (var merge in merges) {
            int rowSpan = Math.Min(merge.RowSpan, rows - merge.Row);
            int columnSpan = Math.Min(merge.ColumnSpan, columns - merge.Column);
            target.MergeCells(merge.Row, merge.Column, rowSpan, columnSpan);
        }
    }

    private static IEnumerable<OdtParagraph> EnumerateTableParagraphs(OdtTable table) {
        foreach (OdtTableRow row in table.Rows) foreach (OdtTableCell cell in row.Cells) {
            foreach (OdtContentBlock block in cell.ReadContentBlocks()) {
                if (block.Paragraph != null) yield return block.Paragraph;
                else foreach (OdtParagraph paragraph in EnumerateTableParagraphs(block.Table!)) yield return paragraph;
            }
        }
    }

    // Work from physical XML runs, before any logical-row enumeration, image
    // inspection or target allocation. Nested tables multiply with their parent cells.
    private static void ValidateOdtTableExpansion(OdtDocument source, WordOpenDocumentConversionOptions options) {
        if (options.MaxTableRows < 1 || options.MaxTableColumns < 1 ||
            options.MaxConvertedTableCells < 1 || options.MaxConvertedTableTextCharacters < 1 || options.MaxTableDepth < 1) {
            throw new ArgumentOutOfRangeException(nameof(options), "Table expansion limits must be positive.");
        }
        long remaining = options.MaxConvertedTableCells;
        long remainingCharacters = options.MaxConvertedTableTextCharacters;
        foreach (OdtContentBlock block in source.ContentBlocks) {
            if (block.Table != null) Inspect(block.Table.Element, 1, 1);
        }

        void Inspect(XElement table, long copies, int depth) {
            if (depth > options.MaxTableDepth) throw new InvalidDataException("ODT table nesting exceeds MaxTableDepth.");
            long rows = 0, columns = 1;
            foreach (XElement row in OdfTableRowElements.Enumerate(table)) {
                long repeats = OdsRepeatModel.Read(row, OdfNamespaces.Table + "number-rows-repeated");
                if (repeats > options.MaxTableRows - rows) throw new InvalidDataException("ODT table exceeds MaxTableRows.");
                rows += repeats;
                long width = 0;
                foreach (XElement cell in Cells(row)) {
                    long cellRepeats = OdsRepeatModel.Read(cell, OdfNamespaces.Table + "number-columns-repeated");
                    if (cellRepeats > options.MaxTableColumns - width) throw new InvalidDataException("ODT table exceeds MaxTableColumns.");
                    width += cellRepeats;
                }
                columns = Math.Max(columns, width);
            }
            rows = Math.Max(1, rows);
            long perCopy = rows * columns;
            if (copies > remaining / perCopy) throw new InvalidDataException("ODT tables exceed MaxConvertedTableCells.");
            remaining -= copies * perCopy;
            foreach (XElement row in OdfTableRowElements.Enumerate(table)) foreach (XElement cell in Cells(row)) {
                long rowRepeats = OdsRepeatModel.Read(row, OdfNamespaces.Table + "number-rows-repeated");
                long cellRepeats = OdsRepeatModel.Read(cell, OdfNamespaces.Table + "number-columns-repeated");
                long occurrences = copies * rowRepeats * cellRepeats;
                foreach (XElement paragraph in cell.Descendants().Where(element =>
                    (element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h") &&
                    element.Ancestors(OdfNamespaces.Table + "table").FirstOrDefault() == table)) {
                    int characters = new OdtParagraph(source, paragraph).Text.Length;
                    if (characters != 0 && occurrences > remainingCharacters / characters) {
                        throw new InvalidDataException("ODT table text exceeds MaxConvertedTableTextCharacters.");
                    }
                    remainingCharacters -= occurrences * characters;
                }
                foreach (XElement nested in cell.Descendants(OdfNamespaces.Table + "table")
                    .Where(candidate => candidate.Ancestors(OdfNamespaces.Table + "table").FirstOrDefault() == table)) {
                    // A nested table allocates at least one cell per copy.
                    if (copies > remaining / rowRepeats / cellRepeats) throw new InvalidDataException("ODT tables exceed MaxConvertedTableCells.");
                    Inspect(nested, copies * rowRepeats * cellRepeats, depth + 1);
                }
            }
        }
        IEnumerable<XElement> Cells(XElement row) => row.Elements().Where(element =>
            element.Name == OdfNamespaces.Table + "table-cell" || element.Name == OdfNamespaces.Table + "covered-table-cell");
    }
}
