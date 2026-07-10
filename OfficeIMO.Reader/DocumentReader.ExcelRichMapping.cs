using OfficeIMO.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Reader;

public static partial class DocumentReader {
    private static OfficeDocumentReadResult ApplyExcelRichMapping(
        ExcelWorkbookSnapshot snapshot,
        ReaderOptions options,
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
        if (selectedSheets.Count == 0 && !string.IsNullOrWhiteSpace(options.ExcelSheetName)) {
            selectedSheets.Add(options.ExcelSheetName!.Trim());
        }

        ExcelWorksheetSnapshot[] worksheets = snapshot.Worksheets
            .Where(sheet => selectedSheets.Count == 0 || selectedSheets.Contains(sheet.Name))
            .ToArray();
        OfficeDocumentLink[] links = BuildExcelLinks(worksheets, result.Source.Path).ToArray();
        ReaderTable[] ownerTables = BuildExcelOwnerTables(worksheets, result.Source.Path, options.MaxTableRows).ToArray();
        IReadOnlyList<ReaderTable> tables = MergeExcelTables(ownerTables, result.Tables);
        IReadOnlyList<OfficeDocumentPage> pages = BuildExcelPages(result.Pages, result.Blocks, tables, result.Assets, links);

        int formulaCount = worksheets.Sum(static sheet => sheet.Cells.Count(static cell => !string.IsNullOrWhiteSpace(cell.Formula)));
        int commentCount = worksheets.Sum(static sheet => sheet.Cells.Count(static cell => cell.Comment != null || cell.ThreadedComment != null));
        var metadata = new List<OfficeDocumentMetadataEntry> {
            BuildCountMetadataEntry("excel-named-range-count", "excel.structure", "NamedRangeCount", snapshot.NamedRanges.Count),
            BuildCountMetadataEntry("excel-formula-count", "excel.structure", "FormulaCount", formulaCount),
            BuildCountMetadataEntry("excel-comment-count", "excel.structure", "CommentCount", commentCount)
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

        return FinalizeRichMapping(
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
        var formalTableSheets = new HashSet<string>(
            ownerTables
                .Where(static table => !string.IsNullOrWhiteSpace(table.Location?.Sheet))
                .Select(static table => table.Location!.Sheet!),
            StringComparer.Ordinal);
        return ownerTables
            .Concat(genericTables.Where(table => string.IsNullOrWhiteSpace(table.Location?.Sheet) || !formalTableSheets.Contains(table.Location!.Sheet!)))
            .ToArray();
    }

    private static IEnumerable<OfficeDocumentLink> BuildExcelLinks(
        IReadOnlyList<ExcelWorksheetSnapshot> worksheets,
        string? sourcePath) {
        int linkIndex = 0;
        foreach (ExcelWorksheetSnapshot sheet in worksheets) {
            foreach (ExcelCellSnapshot cell in sheet.Cells) {
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
        int maxRows) {
        int tableIndex = 0;
        foreach (ExcelWorksheetSnapshot sheet in worksheets) {
            Dictionary<(int Row, int Column), ExcelCellSnapshot> cells = sheet.Cells.ToDictionary(static cell => (cell.Row, cell.Column));
            foreach (ExcelTableSnapshot table in sheet.Tables) {
                int columnCount = Math.Max(table.Columns.Count, table.EndColumn - table.StartColumn + 1);
                IReadOnlyList<string> columns = Enumerable.Range(0, columnCount)
                    .Select(index => index < table.Columns.Count && !string.IsNullOrWhiteSpace(table.Columns[index].Name)
                        ? table.Columns[index].Name
                        : GetExcelCellText(cells, table.StartRow, table.StartColumn + index, "Column " + (index + 1).ToString(CultureInfo.InvariantCulture)))
                    .ToArray();
                int firstDataRow = table.StartRow + (table.HasHeaderRow ? 1 : 0);
                int lastDataRow = table.EndRow - (table.TotalsRowShown ? 1 : 0);
                int totalRows = Math.Max(0, lastDataRow - firstDataRow + 1);
                int emittedRows = maxRows > 0 ? Math.Min(totalRows, maxRows) : totalRows;
                var rows = new List<IReadOnlyList<string>>(emittedRows);
                for (int row = firstDataRow; row < firstDataRow + emittedRows; row++) {
                    rows.Add(Enumerable.Range(0, columnCount)
                        .Select(index => GetExcelCellText(cells, row, table.StartColumn + index, string.Empty))
                        .ToArray());
                }

                yield return new ReaderTable {
                    Title = table.Name,
                    Kind = "excel-table",
                    Location = new ReaderLocation {
                        Path = sourcePath,
                        Sheet = sheet.Name,
                        A1Range = table.A1Range,
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
