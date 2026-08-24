using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    /// <summary>
    /// Helper methods that normalize documents for better rendering in Word Online and Google Docs.
    /// </summary>
    public partial class WordDocument {
        /// <summary>
        /// Walks all tables in body and headers/footers and normalizes their grid/widths
        /// so online viewers render them consistently.
        /// </summary>
        public void NormalizeTablesForOnline() {
            try {
                var main = _wordprocessingDocument.MainDocumentPart;
                if (main == null) return;

                // Body tables
                foreach (var t in main.Document?.Body?.Descendants<Table>() ?? Enumerable.Empty<Table>()) {
                    try {
                        if (IsCanonicalDefaultTable(t)) continue;
                        var wt = new WordTable(this, t, initializeChildren: true);
                        wt.NormalizeForOnline();
                    } catch { }
                }

                // Header tables
                foreach (var hp in main.HeaderParts) {
                    foreach (var t in hp.Header?.Descendants<Table>() ?? Enumerable.Empty<Table>()) {
                        try {
                            if (IsCanonicalDefaultTable(t)) continue;
                            var wt = new WordTable(this, t, initializeChildren: true);
                            wt.NormalizeForOnline();
                        } catch { }
                    }
                }

                // Footer tables
                foreach (var fp in main.FooterParts) {
                    foreach (var t in fp.Footer?.Descendants<Table>() ?? Enumerable.Empty<Table>()) {
                        try {
                            if (IsCanonicalDefaultTable(t)) continue;
                            var wt = new WordTable(this, t, initializeChildren: true);
                            wt.NormalizeForOnline();
                        } catch { }
                    }
                }

            } catch { }
        }

        private static bool IsCanonicalDefaultTable(Table table) {
            if (table.Parent is TableCell ||
                table.Descendants<HorizontalMerge>().Any() ||
                table.Descendants<GridSpan>().Any()) return false;
            TableWidth? tableWidth = table.GetFirstChild<TableProperties>()?.TableWidth;
            if (tableWidth?.Type?.Value != TableWidthUnitValues.Auto || tableWidth.Width?.Value != "0") return false;

            TableRow? firstRow = table.Elements<TableRow>().FirstOrDefault();
            if (firstRow == null) return false;
            int columnCount = firstRow.Elements<TableCell>().Count();
            if (columnCount == 0) return false;
            TableGrid? grid = table.GetFirstChild<TableGrid>();
            if (grid == null || grid.Elements<GridColumn>().Count() != columnCount ||
                grid.Elements<GridColumn>().Any(column => column.Width?.Value != "2400")) return false;

            foreach (TableRow row in table.Elements<TableRow>()) {
                int cellCount = 0;
                foreach (TableCell cell in row.Elements<TableCell>()) {
                    cellCount++;
                    TableCellWidth? width = cell.TableCellProperties?.TableCellWidth;
                    if (width?.Type?.Value != TableWidthUnitValues.Dxa || width.Width?.Value != "2400") return false;
                }
                if (cellCount != columnCount) return false;
            }
            return true;
        }
    }
}
