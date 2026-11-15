using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    // Helper methods that normalize documents for better rendering in Word Online and Google Docs.
    public partial class WordDocument {
        /// <summary>
        /// Walks all tables in body and headers/footers and normalizes their grid/widths
        /// so online viewers render them consistently.
        /// </summary>
        public void NormalizeTablesForOnline() => NormalizeTablesForOnline(knownTables: null);

        /// <param name="knownTables">Tables already collected, in document order, for part roots that have not
        /// changed since; the main document's entry is keyed by its <see cref="Document"/>.</param>
        private void NormalizeTablesForOnline(Dictionary<OpenXmlElement, List<Table>>? knownTables) {
            if (!_tableNormalizationRequired) return;

            try {
                var main = _wordprocessingDocument.MainDocumentPart;
                if (main == null) return;

                Document? document = main.Document;
                // The document root holds tables only inside its body.
                NormalizeTablesForOnline(document != null && knownTables != null && knownTables.TryGetValue(document, out List<Table>? bodyTables)
                    ? bodyTables
                    : CollectTables(document?.Body));
                foreach (var hp in main.HeaderParts) NormalizeTablesForOnline(KnownOrCollectedTables(hp.Header, knownTables));
                foreach (var fp in main.FooterParts) NormalizeTablesForOnline(KnownOrCollectedTables(fp.Footer, knownTables));
            } catch { }
        }

        private static List<Table> KnownOrCollectedTables(OpenXmlElement? root, Dictionary<OpenXmlElement, List<Table>>? knownTables) =>
            root != null && knownTables != null && knownTables.TryGetValue(root, out List<Table>? tables) ? tables : CollectTables(root);

        /// <summary>Tables in document order, as <c>Descendants&lt;Table&gt;()</c> returns them.</summary>
        private static List<Table> CollectTables(OpenXmlElement? root) {
            var tables = new List<Table>();
            if (root == null) return tables;
            for (OpenXmlElement? element = root.FirstChild; element != null;
                 element = WordElementWalk.Next(element, root, descend: !WordElementWalk.IsFormattingOrText(element))) {
                if (!WordElementWalk.IsCommonContent(element) && element is Table table) tables.Add(table);
            }
            return tables;
        }

        private void NormalizeTablesForOnline(List<Table> tables) {
            foreach (Table t in tables) {
                try {
                    if (IsCanonicalDefaultTable(t)) continue;
                    var wt = new WordTable(this, t, initializeChildren: true);
                    wt.NormalizeForOnline();
                } catch { }
            }
        }

        internal bool TableNormalizationRequired => _tableNormalizationRequired;

        internal void MarkTableNormalizationRequired() => _tableNormalizationRequired = true;

        private static bool IsCanonicalDefaultTable(Table table) {
            if (table.Parent is TableCell || HasMergedCells(table)) return false;
            TableWidth? tableWidth = table.GetFirstChild<TableProperties>()?.TableWidth;
            if (tableWidth?.Type?.Value != TableWidthUnitValues.Auto || tableWidth.Width?.Value != "0") return false;

            TableRow? firstRow = table.Elements<TableRow>().FirstOrDefault();
            if (firstRow == null) return false;
            int columnCount = firstRow.Elements<TableCell>().Count();
            if (columnCount == 0) return false;
            TableGrid? grid = table.GetFirstChild<TableGrid>();
            // A freshly generated table has one uniform column width on its grid and cells.
            string? uniformWidth = grid?.Elements<GridColumn>().FirstOrDefault()?.Width?.Value;
            if (grid == null || uniformWidth == null || grid.Elements<GridColumn>().Count() != columnCount ||
                grid.Elements<GridColumn>().Any(column => column.Width?.Value != uniformWidth)) return false;

            // Sibling loops rather than Elements<T>(), which allocates an enumerator for every row.
            for (OpenXmlElement? rowElement = table.FirstChild; rowElement != null; rowElement = rowElement.NextSibling()) {
                if (rowElement is not TableRow row) continue;
                int cellCount = 0;
                for (OpenXmlElement? cellElement = row.FirstChild; cellElement != null; cellElement = cellElement.NextSibling()) {
                    if (cellElement is not TableCell cell) continue;
                    cellCount++;
                    // Plain child scans; the schema-driven property lookups cost more for every cell.
                    TableCellWidth? width = WordElementWalk.FirstChild<TableCellProperties>(cell) is { } cellProperties ? WordElementWalk.FirstChild<TableCellWidth>(cellProperties) : null;
                    if (width?.Type?.Value != TableWidthUnitValues.Dxa || width.Width?.Value != uniformWidth) return false;
                }
                if (cellCount != columnCount) return false;
            }
            return true;
        }

        private static bool HasMergedCells(Table table) {
            for (OpenXmlElement? element = table.FirstChild; element != null;
                 element = WordElementWalk.Next(element, table, descend: !WordElementWalk.IsFormattingOrText(element))) {
                if (!WordElementWalk.IsCommonContent(element) && element is HorizontalMerge or GridSpan) return true;
            }
            return false;
        }
    }
}
