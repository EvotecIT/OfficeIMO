using OfficeIMO.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Reader.Excel;

internal static class ExcelRichMapping {
    internal static OfficeDocumentReadResult Apply(
        ExcelWorkbookSnapshot snapshot,
        ReaderOptions readerOptions,
        ReaderExcelOptions options,
        OfficeDocumentReadResult result) {
        result.Source.Title = snapshot.Title;
        result.Source.Author = snapshot.Author;
        result.Source.Subject = snapshot.Subject;
        result.Source.Keywords = snapshot.Keywords;

        var selectedSheets = new HashSet<string>(
            result.Pages
                .Where(static page => !string.IsNullOrWhiteSpace(page.Name))
                .Select(static page => page.Name!),
            StringComparer.Ordinal);
        if (selectedSheets.Count == 0 && !string.IsNullOrWhiteSpace(options.SheetName)) {
            selectedSheets.Add(options.SheetName!.Trim());
        }

        ExcelWorksheetSnapshot[] worksheets = snapshot.Worksheets
            .Where(sheet => selectedSheets.Count == 0 || selectedSheets.Contains(sheet.Name))
            .ToArray();
        ExcelRangeSelection? selectedRange = ParseExcelRangeSelection(options.A1Range);
        OfficeDocumentLink[] links = BuildExcelLinks(worksheets, result.Source.Path, selectedRange).ToArray();
        ReaderTable[] ownerTables = BuildExcelOwnerTables(worksheets, result.Source.Path, readerOptions.MaxTableRows, selectedRange).ToArray();
        IReadOnlyList<ReaderTable> tables = MergeExcelTables(ownerTables, result.Tables);
        IReadOnlyList<OfficeDocumentPage> pages = BuildExcelPages(result.Pages, result.Blocks, tables, result.Assets, links);

        int formulaCount = worksheets.Sum(sheet => sheet.Cells.Count(cell => IsCellSelected(cell, selectedRange) && !string.IsNullOrWhiteSpace(cell.Formula)));
        int commentCount = worksheets.Sum(sheet => sheet.Cells.Count(cell => IsCellSelected(cell, selectedRange) && (cell.Comment != null || cell.ThreadedComment != null)));
        var metadata = new List<OfficeDocumentMetadataEntry> {
            DocumentReaderEngine.BuildCountMetadataEntry("excel-named-range-count", "excel.structure", "NamedRangeCount", snapshot.NamedRanges.Count),
            DocumentReaderEngine.BuildCountMetadataEntry("excel-formula-count", "excel.structure", "FormulaCount", formulaCount),
            DocumentReaderEngine.BuildCountMetadataEntry("excel-comment-count", "excel.structure", "CommentCount", commentCount)
        };
        if (!string.IsNullOrWhiteSpace(snapshot.ActiveWorksheetName)) {
            metadata.Add(new OfficeDocumentMetadataEntry {
                Id = "excel-active-worksheet",
                Category = "excel.workbook",
                Name = "ActiveWorksheet",
                Value = snapshot.ActiveWorksheetName,
                ValueType = "string"
            });
        }

        return DocumentReaderEngine.EnrichDocumentResult(
            result,
            new[] { "officeimo.excel.inspection-snapshot", "officeimo.reader.excel.rich-v5" },
            result.Blocks,
            tables,
            links,
            result.Visuals,
            pages,
            metadata);
    }

    private static IReadOnlyList<ReaderTable> MergeExcelTables(
        IReadOnlyList<ReaderTable> ownerTables,
        IReadOnlyList<ReaderTable> genericTables) {
        if (ownerTables.Count == 0) return genericTables;
        return ownerTables
            .Concat(genericTables.Where(generic => !ownerTables.Any(owner => ExcelTableRangesOverlap(owner, generic))))
            .ToArray();
    }

    private static bool ExcelTableRangesOverlap(ReaderTable first, ReaderTable second) {
        if (!string.Equals(first.Location?.Sheet, second.Location?.Sheet, StringComparison.Ordinal)
            || string.IsNullOrWhiteSpace(first.Location?.A1Range)
            || string.IsNullOrWhiteSpace(second.Location?.A1Range)) {
            return false;
        }

        ExcelRangeSelection? firstRange = ParseExcelRangeSelection(first.Location!.A1Range);
        ExcelRangeSelection? secondRange = ParseExcelRangeSelection(second.Location!.A1Range);
        return firstRange.HasValue
            && secondRange.HasValue
            && firstRange.Value.StartRow <= secondRange.Value.EndRow
            && secondRange.Value.StartRow <= firstRange.Value.EndRow
            && firstRange.Value.StartColumn <= secondRange.Value.EndColumn
            && secondRange.Value.StartColumn <= firstRange.Value.EndColumn;
    }

    private static IEnumerable<OfficeDocumentLink> BuildExcelLinks(
        IReadOnlyList<ExcelWorksheetSnapshot> worksheets,
        string? sourcePath,
        ExcelRangeSelection? selectedRange) {
        int linkIndex = 0;
        foreach (ExcelWorksheetSnapshot sheet in worksheets) {
            foreach (ExcelCellSnapshot cell in sheet.Cells) {
                if (!IsCellSelected(cell, selectedRange)) continue;
                ExcelHyperlinkSnapshot? hyperlink = cell.Hyperlink;
                if (hyperlink == null || string.IsNullOrWhiteSpace(hyperlink.Target)) continue;
                string a1 = A1.CellReference(cell.Row, cell.Column);
                yield return new OfficeDocumentLink {
                    Id = "excel-link-" + linkIndex.ToString("D4", CultureInfo.InvariantCulture),
                    Kind = hyperlink.IsExternal ? "uri" : "internal",
                    Uri = hyperlink.IsExternal ? hyperlink.Target : null,
                    DestinationName = hyperlink.IsExternal ? null : hyperlink.Target,
                    Text = hyperlink.Tooltip ?? Convert.ToString(cell.Value, CultureInfo.InvariantCulture),
                    Location = new ReaderLocation {
                        Path = sourcePath,
                        Sheet = sheet.Name,
                        A1Range = a1,
                        SourceBlockKind = "hyperlink",
                        BlockAnchor = "excel-" + SanitizeAnchor(sheet.Name) + "-" + a1.ToLowerInvariant() + "-link"
                    }
                };
                linkIndex++;
            }
        }
    }

    private static IEnumerable<ReaderTable> BuildExcelOwnerTables(
        IReadOnlyList<ExcelWorksheetSnapshot> worksheets,
        string? sourcePath,
        int maxRows,
        ExcelRangeSelection? selectedRange) {
        int tableIndex = 0;
        foreach (ExcelWorksheetSnapshot sheet in worksheets) {
            Dictionary<(int Row, int Column), ExcelCellSnapshot> cells = sheet.Cells.ToDictionary(static cell => (cell.Row, cell.Column));
            foreach (ExcelTableSnapshot table in sheet.Tables) {
                ExcelRangeSelection tableRange = new ExcelRangeSelection(table.StartRow, table.StartColumn, table.EndRow, table.EndColumn);
                ExcelRangeSelection? projectedRange = IntersectExcelRanges(tableRange, selectedRange);
                if (!projectedRange.HasValue) continue;

                int startColumn = projectedRange.Value.StartColumn;
                int endColumn = projectedRange.Value.EndColumn;
                int columnCount = endColumn - startColumn + 1;
                IReadOnlyList<string> columns = Enumerable.Range(0, columnCount)
                    .Select(index => {
                        int tableColumnIndex = startColumn + index - table.StartColumn;
                        return tableColumnIndex >= 0
                            && tableColumnIndex < table.Columns.Count
                            && !string.IsNullOrWhiteSpace(table.Columns[tableColumnIndex].Name)
                                ? table.Columns[tableColumnIndex].Name
                                : GetExcelCellText(cells, table.StartRow, startColumn + index, "Column " + (index + 1).ToString(CultureInfo.InvariantCulture));
                    })
                    .ToArray();
                int firstDataRow = Math.Max(projectedRange.Value.StartRow, table.StartRow + (table.HasHeaderRow ? 1 : 0));
                int lastDataRow = Math.Min(projectedRange.Value.EndRow, table.EndRow - (table.TotalsRowShown ? 1 : 0));
                int totalRows = Math.Max(0, lastDataRow - firstDataRow + 1);
                int emittedRows = maxRows > 0 ? Math.Min(totalRows, maxRows) : totalRows;
                var rows = new List<IReadOnlyList<string>>(emittedRows);
                for (int row = firstDataRow; row < firstDataRow + emittedRows; row++) {
                    rows.Add(Enumerable.Range(0, columnCount)
                        .Select(index => GetExcelCellText(cells, row, startColumn + index, string.Empty))
                        .ToArray());
                }

                yield return new ReaderTable {
                    Title = table.Name,
                    Kind = "excel-table",
                    Location = new ReaderLocation {
                        Path = sourcePath,
                        Sheet = sheet.Name,
                        A1Range = projectedRange.Value.ToA1Range(),
                        SourceBlockKind = "table",
                        BlockAnchor = "excel-" + SanitizeAnchor(sheet.Name) + "-table-" + tableIndex.ToString("D4", CultureInfo.InvariantCulture),
                        TableIndex = tableIndex
                    },
                    Columns = columns,
                    ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, rows),
                    Rows = rows,
                    TotalRowCount = totalRows,
                    Truncated = emittedRows < totalRows
                };
                tableIndex++;
            }
        }
    }

    private static ExcelRangeSelection? ParseExcelRangeSelection(string? a1Range) {
        if (string.IsNullOrWhiteSpace(a1Range)
            || !A1.TryParseRange(a1Range!, out int startRow, out int startColumn, out int endRow, out int endColumn)) {
            return null;
        }
        return new ExcelRangeSelection(startRow, startColumn, endRow, endColumn);
    }

    private static ExcelRangeSelection? IntersectExcelRanges(ExcelRangeSelection range, ExcelRangeSelection? selection) {
        if (!selection.HasValue) return range;
        int startRow = Math.Max(range.StartRow, selection.Value.StartRow);
        int startColumn = Math.Max(range.StartColumn, selection.Value.StartColumn);
        int endRow = Math.Min(range.EndRow, selection.Value.EndRow);
        int endColumn = Math.Min(range.EndColumn, selection.Value.EndColumn);
        return startRow > endRow || startColumn > endColumn
            ? (ExcelRangeSelection?)null
            : new ExcelRangeSelection(startRow, startColumn, endRow, endColumn);
    }

    private static bool IsCellSelected(ExcelCellSnapshot cell, ExcelRangeSelection? selection) {
        return !selection.HasValue
            || (cell.Row >= selection.Value.StartRow
                && cell.Row <= selection.Value.EndRow
                && cell.Column >= selection.Value.StartColumn
                && cell.Column <= selection.Value.EndColumn);
    }

    private readonly struct ExcelRangeSelection : IEquatable<ExcelRangeSelection> {
        internal ExcelRangeSelection(int startRow, int startColumn, int endRow, int endColumn) {
            StartRow = startRow;
            StartColumn = startColumn;
            EndRow = endRow;
            EndColumn = endColumn;
        }

        internal int StartRow { get; }
        internal int StartColumn { get; }
        internal int EndRow { get; }
        internal int EndColumn { get; }

        internal string ToA1Range() {
            return A1.CellReference(StartRow, StartColumn) + ":" + A1.CellReference(EndRow, EndColumn);
        }

        public bool Equals(ExcelRangeSelection other) {
            return StartRow == other.StartRow
                && StartColumn == other.StartColumn
                && EndRow == other.EndRow
                && EndColumn == other.EndColumn;
        }

        public override bool Equals(object? obj) {
            return obj is ExcelRangeSelection other && Equals(other);
        }

        public override int GetHashCode() {
            unchecked {
                int hash = StartRow;
                hash = (hash * 397) ^ StartColumn;
                hash = (hash * 397) ^ EndRow;
                return (hash * 397) ^ EndColumn;
            }
        }
    }

    private static string GetExcelCellText(
        IReadOnlyDictionary<(int Row, int Column), ExcelCellSnapshot> cells,
        int row,
        int column,
        string fallback) {
        if (!cells.TryGetValue((row, column), out ExcelCellSnapshot? cell) || cell.Value == null) return fallback;
        string? value = Convert.ToString(cell.Value, CultureInfo.InvariantCulture);
        return string.IsNullOrWhiteSpace(value) ? fallback : value!;
    }

    private static IReadOnlyList<OfficeDocumentPage> BuildExcelPages(
        IReadOnlyList<OfficeDocumentPage> existingPages,
        IReadOnlyList<OfficeDocumentBlock> blocks,
        IReadOnlyList<ReaderTable> tables,
        IReadOnlyList<OfficeDocumentAsset> assets,
        IReadOnlyList<OfficeDocumentLink> links) {
        foreach (OfficeDocumentPage page in existingPages) {
            string? sheet = page.Name ?? page.Location.Sheet;
            page.Blocks = blocks.Where(block => string.Equals(block.Location.Sheet, sheet, StringComparison.Ordinal)).ToArray();
            page.Tables = tables.Where(table => string.Equals(table.Location?.Sheet, sheet, StringComparison.Ordinal)).ToArray();
            page.Assets = assets.Where(asset => string.Equals(asset.Location.Sheet, sheet, StringComparison.Ordinal)).ToArray();
            page.Links = links.Where(link => string.Equals(link.Location.Sheet, sheet, StringComparison.Ordinal)).ToArray();
        }
        return existingPages;
    }

    private static string SanitizeAnchor(string value) {
        if (string.IsNullOrWhiteSpace(value)) return "item";
        var chars = value.Trim().ToLowerInvariant().Select(static ch => char.IsLetterOrDigit(ch) ? ch : '-').ToArray();
        return new string(chars).Trim('-');
    }
}
