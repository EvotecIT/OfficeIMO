using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Sets an external hyperlink on a single cell. If display text is null or empty, the URL is shown.
        /// </summary>
        public void SetHyperlink(int row, int column, string url, string? display = null) {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));

            WriteLock(() => {
                var cell = GetCell(row, column);
                var text = string.IsNullOrEmpty(display) ? url : display;
                // Avoid nested locks: write value using core method
                CellValueCore(row, column, text);

                // Ensure Hyperlinks container exists
                var ws = _worksheetPart.Worksheet;
                var hyperlinks = ws.Elements<Hyperlinks>().FirstOrDefault();
                if (hyperlinks == null) {
                    hyperlinks = new Hyperlinks();
                    // place near the end but before TableParts per schema order
                    var tableParts = ws.Elements<TableParts>().FirstOrDefault();
                    if (tableParts != null) ws.InsertBefore(hyperlinks, tableParts); else ws.Append(hyperlinks);
                }

                // Add external relationship
                var rel = _worksheetPart.AddHyperlinkRelationship(new Uri(url), true);
                var reference = GetColumnName(column) + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
                var hl = new Hyperlink { Reference = reference, Id = rel.Id };
                hyperlinks.Append(hl);
            });
        }

        /// <summary>
        /// Sets an external hyperlink using an A1 reference (e.g., "B5").
        /// </summary>
        public void SetHyperlink(string a1, string url, string? display = null) {
            var col = GetColumnIndex(a1);
            var row = GetRowIndex(a1);
            SetHyperlink(row, col, url, display);
        }

        /// <summary>
        /// Sets an internal hyperlink (location in this workbook), e.g., "'Sheet1'!A1".
        /// </summary>
        public void SetInternalLink(int row, int column, string location, string? display = null)
        {
            if (string.IsNullOrWhiteSpace(location)) throw new ArgumentNullException(nameof(location));
            WriteLock(() =>
            {
                var text = string.IsNullOrEmpty(display) ? location : display;
                CellValueCore(row, column, text);
                var ws = _worksheetPart.Worksheet;
                var hyperlinks = ws.Elements<Hyperlinks>().FirstOrDefault();
                if (hyperlinks == null)
                {
                    hyperlinks = new Hyperlinks();
                    var tableParts = ws.Elements<TableParts>().FirstOrDefault();
                    if (tableParts != null) ws.InsertBefore(hyperlinks, tableParts); else ws.Append(hyperlinks);
                }
                var reference = GetColumnName(column) + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
                var hl = new Hyperlink { Reference = reference, Location = location };
                hyperlinks.Append(hl);
                // Defer save to caller; final document Save() will persist
            });
        }
    }
}
