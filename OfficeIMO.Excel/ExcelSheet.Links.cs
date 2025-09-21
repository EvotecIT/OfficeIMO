using System;
using System.Text.RegularExpressions;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel
{
    public partial class ExcelSheet
    {
        /// <summary>
        /// Parses an A1 range and returns 1-based bounds (r1, c1, r2, c2).
        /// </summary>
        /// <param name="a1">A1 range without a sheet prefix (e.g., "A2:D20").</param>
        /// <returns>Tuple of (r1, c1, r2, c2) with normalized bounds.</returns>
        /// <exception cref="ArgumentException">Thrown when the input is not a valid A1 range.</exception>
        public (int r1, int c1, int r2, int c2) GetRangeBounds(string a1)
        {
            return A1.ParseRange(a1);
        }
        /// <example>
        /// var bounds = sheet.GetRangeBounds("A2:D20");
        /// // bounds.r1=2, bounds.c1=1, bounds.r2=20, bounds.c2=4
        /// </example>

        /// <summary>
        /// Iterates over row indices inside an A1 range and invokes <paramref name="action"/> for each row.
        /// </summary>
        /// <param name="a1">A1 range without a sheet prefix.</param>
        /// <param name="action">Callback receiving a 1-based row index.</param>
        public void ForEachRow(string a1, Action<int> action)
        {
            var (r1, _, r2, _) = A1.ParseRange(a1);
            for (int r = r1; r <= r2; r++) action(r);
        }
        /// <example>
        /// sheet.ForEachRow("A2:A10", r => sheet.SetInternalLink(r, 1, "'Summary'!A1", "Back"));
        /// </example>

        /// <summary>
        /// Iterates over column indices inside an A1 range and invokes <paramref name="action"/> for each column.
        /// </summary>
        /// <param name="a1">A1 range without a sheet prefix.</param>
        /// <param name="action">Callback receiving a 1-based column index.</param>
        public void ForEachColumn(string a1, Action<int> action)
        {
            var (_, c1, _, c2) = A1.ParseRange(a1);
            for (int c = c1; c <= c2; c++) action(c);
        }
        /// <example>
        /// sheet.ForEachColumn("A1:E1", c => sheet.CellBold(1, c, true));
        /// </example>

        /// <summary>
        /// Converts each non-empty cell in the A1 range into an internal hyperlink.
        /// The destination sheet name is computed from the cell's text using <paramref name="destinationSheetForCellText"/>.
        /// </summary>
        /// <param name="a1">A1 range to process (e.g., a column of names).</param>
        /// <param name="destinationSheetForCellText">Maps the cell text to a destination sheet name.</param>
        /// <param name="targetA1">Destination cell on the target sheet (default "A1").</param>
        /// <param name="display">Optional display text selector. Defaults to the cell text.</param>
        /// <param name="styled">When true, applies hyperlink styling (blue + underline).</param>
        public void LinkCellsToInternalSheets(string a1, Func<string, string> destinationSheetForCellText, string targetA1 = "A1", Func<string, string>? display = null, bool styled = true)
        {
            if (destinationSheetForCellText == null) throw new ArgumentNullException(nameof(destinationSheetForCellText));
            var (r1, c1, r2, c2) = A1.ParseRange(a1);
            for (int r = r1; r <= r2; r++)
            {
                for (int c = c1; c <= c2; c++)
                {
                    if (!TryGetCellText(r, c, out var text) || string.IsNullOrWhiteSpace(text)) continue;
                    string sheetName = destinationSheetForCellText(text);
                    if (string.IsNullOrWhiteSpace(sheetName)) continue;
                    string location = $"'{EscapeSheetNameForLink(sheetName)}'!{targetA1}";
                    string disp = display?.Invoke(text) ?? text;
                    SetInternalLink(r, c, location, disp, styled);
                }
            }
        }
        /// <example>
        /// // Given a summary table where column A contains sheet names, link each cell to its sheet
        /// sheet.LinkCellsToInternalSheets("A2:A51", text => text, targetA1: "A1", styled: true);
        /// </example>

        /// <summary>
        /// Creates an external hyperlink using a smart display strategy: prefer <paramref name="title"/>,
        /// then an RFC label (e.g., "RFC 7208") when detected, otherwise the URL host.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="url">Target URL.</param>
        /// <param name="title">Optional preferred display text.</param>
        /// <param name="style">When true, applies hyperlink styling (blue + underline).</param>
        public void SetHyperlinkSmart(int row, int column, string url, string? title = null, bool style = true)
        {
            string display = !string.IsNullOrWhiteSpace(title) ? title! : GuessLinkDisplay(url);
            SetHyperlink(row, column, url, display, style);
        }
        /// <example>
        /// sheet.SetHyperlinkSmart(5, 1, "https://datatracker.ietf.org/doc/html/rfc7208"); // displays "RFC 7208"
        /// sheet.SetHyperlinkSmart(6, 1, "https://example.org/path", title: "Spec");    // displays "Spec"
        /// </example>

        /// <summary>
        /// Creates an external hyperlink showing only the host (e.g., example.org) as display text.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="url">Target URL.</param>
        /// <param name="style">When true, applies hyperlink styling (blue + underline).</param>
        public void SetHyperlinkHost(int row, int column, string url, bool style = true)
        {
            string display = TryGetHost(url, out var host) ? host! : url;
            SetHyperlink(row, column, url, display, style);
        }
        /// <example>
        /// sheet.SetHyperlinkHost(7, 1, "https://learn.microsoft.com/office/open-xml/"); // displays "learn.microsoft.com"
        /// </example>

        private static string GuessLinkDisplay(string url)
        {
            // RFC pattern: https://datatracker.ietf.org/doc/html/rfc7208
            var r = Regex.Match(url ?? string.Empty, @"rfc(\d{3,5})", RegexOptions.IgnoreCase);
            if (r.Success) return "RFC " + r.Groups[1].Value;
            if (TryGetHost(url, out var host)) return host!;
            return url ?? string.Empty;
        }

        private static bool TryGetHost(string? url, out string? host)
        {
            host = null;
            if (string.IsNullOrWhiteSpace(url)) return false;
            if (Uri.TryCreate(url, UriKind.Absolute, out var uri))
            {
                host = uri.Host;
                return !string.IsNullOrEmpty(host);
            }
            return false;
        }

        /// <summary>
        /// Links a column identified by <paramref name="header"/> to internal sheets using each cell's text.
        /// Defaults to linking to a sheet with the same name as the cell text.
        /// </summary>
        /// <param name="header">Header text of the column to process.</param>
        /// <param name="rowFrom">First data row (1-based). Defaults to 2 (skip header).</param>
        /// <param name="rowTo">Last data row (inclusive). When &lt;= 0, uses the bottom of the used range.</param>
        /// <param name="destinationSheetForCellText">Maps cell text to a destination sheet name; defaults to identity.</param>
        /// <param name="targetA1">Destination cell on the target sheet (default "A1").</param>
        /// <param name="display">Optional display selector; defaults to the cell text.</param>
        /// <param name="styled">Apply hyperlink styling (blue + underline).</param>
        public void LinkByHeaderToInternalSheets(
            string header,
            int rowFrom = 2,
            int rowTo = -1,
            Func<string, string>? destinationSheetForCellText = null,
            string targetA1 = "A1",
            Func<string, string>? display = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            int col = ColumnIndexByHeader(header);
            if (rowTo <= 0)
            {
                var (r1, _, r2, _) = A1.ParseRange(GetUsedRangeA1());
                rowTo = r2;
                if (rowFrom < r1) rowFrom = r1 + 1; // move below header
            }
            var toSheet = destinationSheetForCellText ?? (text => text);
            for (int r = rowFrom; r <= rowTo; r++)
            {
                if (!TryGetCellText(r, col, out var text) || string.IsNullOrWhiteSpace(text)) continue;
                string sheetName = toSheet(text);
                if (string.IsNullOrWhiteSpace(sheetName)) continue;
                string location = $"'{EscapeSheetNameForLink(sheetName)}'!{targetA1}";
                string disp = display?.Invoke(text) ?? text;
                SetInternalLink(r, col, location, disp, styled);
            }
        }

        /// <summary>
        /// Non-throwing variant of <see cref="LinkByHeaderToInternalSheets"/>.
        /// Returns false when the header cannot be found or inputs are invalid.
        /// </summary>
        public bool TryLinkByHeaderToInternalSheets(
            string header,
            int rowFrom = 2,
            int rowTo = -1,
            Func<string, string>? destinationSheetForCellText = null,
            string targetA1 = "A1",
            Func<string, string>? display = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(header)) return false;
            if (!TryGetColumnIndexByHeader(header, out var col)) return false;
            if (rowTo <= 0)
            {
                var ok = A1.TryParseRange(GetUsedRangeA1(), out var r1, out _, out var r2, out _);
                if (!ok) return false;
                rowTo = r2;
                if (rowFrom < r1) rowFrom = r1 + 1;
            }
            var toSheet = destinationSheetForCellText ?? (text => text);
            for (int r = rowFrom; r <= rowTo; r++)
            {
                if (!TryGetCellText(r, col, out var text) || string.IsNullOrWhiteSpace(text)) continue;
                string sheetName = toSheet(text);
                if (string.IsNullOrWhiteSpace(sheetName)) continue;
                string location = $"'{EscapeSheetNameForLink(sheetName)}'!{targetA1}";
                string disp = display?.Invoke(text) ?? text;
                SetInternalLink(r, col, location, disp, styled);
            }
            return true;
        }

        /// <summary>
        /// Links a column identified by <paramref name="header"/> to external URLs built from each cell's text.
        /// </summary>
        /// <param name="header">Header text of the column to process.</param>
        /// <param name="rowFrom">First data row (1-based). Defaults to 2 (skip header).</param>
        /// <param name="rowTo">Last data row (inclusive). When &lt;= 0, uses the bottom of the used range.</param>
        /// <param name="urlForCellText">Maps cell text to URL.</param>
        /// <param name="titleForCellText">Optional display selector; when null, a smart display (RFC/host) is used.</param>
        /// <param name="styled">Apply hyperlink styling (blue + underline).</param>
        public void LinkByHeaderToUrls(
            string header,
            int rowFrom = 2,
            int rowTo = -1,
            Func<string, string> urlForCellText = null!,
            Func<string, string>? titleForCellText = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (urlForCellText is null) throw new ArgumentNullException(nameof(urlForCellText));
            int col = ColumnIndexByHeader(header);
            if (rowTo <= 0)
            {
                var (r1, _, r2, _) = A1.ParseRange(GetUsedRangeA1());
                rowTo = r2;
                if (rowFrom < r1) rowFrom = r1 + 1;
            }
            for (int r = rowFrom; r <= rowTo; r++)
            {
                if (!TryGetCellText(r, col, out var text) || string.IsNullOrWhiteSpace(text)) continue;
                string url = urlForCellText(text);
                if (string.IsNullOrWhiteSpace(url)) continue;
                if (titleForCellText != null)
                    SetHyperlink(r, col, url, titleForCellText(text), styled);
                else
                    SetHyperlinkSmart(r, col, url, null, styled);
            }
        }

        /// <summary>
        /// Non-throwing variant of <see cref="LinkByHeaderToUrls"/>.
        /// Returns false when the header cannot be found or inputs are invalid.
        /// </summary>
        public bool TryLinkByHeaderToUrls(
            string header,
            int rowFrom = 2,
            int rowTo = -1,
            Func<string, string>? urlForCellText = null,
            Func<string, string>? titleForCellText = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(header)) return false;
            if (urlForCellText is null) return false;
            if (!TryGetColumnIndexByHeader(header, out var col)) return false;
            if (rowTo <= 0)
            {
                var ok = A1.TryParseRange(GetUsedRangeA1(), out var r1, out _, out var r2, out _);
                if (!ok) return false;
                rowTo = r2;
                if (rowFrom < r1) rowFrom = r1 + 1;
            }
            for (int r = rowFrom; r <= rowTo; r++)
            {
                if (!TryGetCellText(r, col, out var text) || string.IsNullOrWhiteSpace(text)) continue;
                string url = urlForCellText(text);
                if (string.IsNullOrWhiteSpace(url)) continue;
                if (titleForCellText != null)
                    SetHyperlink(r, col, url, titleForCellText(text), styled);
                else
                    SetHyperlinkSmart(r, col, url, null, styled);
            }
            return true;
        }

        /// <example>
        /// // Internal: link "Domain" column to same-named sheets, rows 2..used bottom
        /// sheet.LinkByHeaderToInternalSheets("Domain");
        /// // External: link "RFC" column to IETF datatracker pages
        /// sheet.LinkByHeaderToUrls("RFC", urlForCellText: rfc => $"https://datatracker.ietf.org/doc/html/{rfc}");
        /// </example>

        /// <summary>
        /// Links a column within a table to internal sheets. The table range determines the row bounds.
        /// </summary>
        /// <param name="tableName">Name of the table (as shown in Excel's Name Manager).</param>
        /// <param name="header">Header text of the column inside the table.</param>
        /// <param name="destinationSheetForCellText">Maps cell text to a destination sheet name (defaults to identity).</param>
        /// <param name="targetA1">Destination cell on the target sheet (default "A1").</param>
        /// <param name="display">Optional display selector; defaults to the cell text.</param>
        /// <param name="styled">Apply hyperlink styling (blue + underline).</param>
        public void LinkByHeaderToInternalSheetsInTable(
            string tableName,
            string header,
            Func<string, string>? destinationSheetForCellText = null,
            string targetA1 = "A1",
            Func<string, string>? display = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentNullException(nameof(tableName));
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            var tdp = _worksheetPart.TableDefinitionParts.FirstOrDefault(p => string.Equals(p.Table?.Name?.Value, tableName, StringComparison.OrdinalIgnoreCase));
            if (tdp?.Table == null) throw new InvalidOperationException($"Table '{tableName}' not found on sheet '{Name}'.");
            var table = tdp.Table;
            string? refA1 = table.Reference?.Value;
            if (string.IsNullOrWhiteSpace(refA1)) throw new InvalidOperationException($"Table '{tableName}' has no Reference.");
            var (r1, c1, r2, c2) = A1.ParseRange(refA1!);
            // Determine header offset
            var headers = table.TableColumns?.Elements<DocumentFormat.OpenXml.Spreadsheet.TableColumn>().Select(tc => tc.Name?.Value ?? string.Empty).ToList() ?? new System.Collections.Generic.List<string>();
            int colOffset = headers.FindIndex(h => string.Equals(h, header, StringComparison.OrdinalIgnoreCase));
            if (colOffset < 0) throw new InvalidOperationException($"Header '{header}' not found in table '{tableName}'.");
            // Header and totals
            int headerRows = (int)(table.HeaderRowCount?.Value ?? 1U);
            bool totals = table.TotalsRowShown?.Value ?? false;
            int totalsRows = totals ? 1 : 0;
            int startRow = r1 + headerRows;
            int endRow = r2 - totalsRows;
            int column = c1 + colOffset;
            var toSheet = destinationSheetForCellText ?? (text => text);
            for (int r = startRow; r <= endRow; r++)
            {
                if (!TryGetCellText(r, column, out var text) || string.IsNullOrWhiteSpace(text)) continue;
                string sheetName = toSheet(text);
                if (string.IsNullOrWhiteSpace(sheetName)) continue;
                string location = $"'{EscapeSheetNameForLink(sheetName)}'!{targetA1}";
                string disp = display?.Invoke(text) ?? text;
                SetInternalLink(r, column, location, disp, styled);
            }
        }

        /// <summary>
        /// Non-throwing variant of <see cref="LinkByHeaderToInternalSheetsInTable"/>.
        /// Returns false when the table or header cannot be found.
        /// </summary>
        public bool TryLinkByHeaderToInternalSheetsInTable(
            string tableName,
            string header,
            Func<string, string>? destinationSheetForCellText = null,
            string targetA1 = "A1",
            Func<string, string>? display = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(tableName) || string.IsNullOrWhiteSpace(header)) return false;
            var tdp = _worksheetPart.TableDefinitionParts.FirstOrDefault(p => string.Equals(p.Table?.Name?.Value, tableName, StringComparison.OrdinalIgnoreCase));
            var table = tdp?.Table; if (table == null) return false;
            string? refA1 = table.Reference?.Value; if (string.IsNullOrWhiteSpace(refA1)) return false;
            if (!A1.TryParseRange(refA1!, out var r1, out var c1, out var r2, out var c2)) return false;
            var headers = table.TableColumns?.Elements<DocumentFormat.OpenXml.Spreadsheet.TableColumn>().Select(tc => tc.Name?.Value ?? string.Empty).ToList() ?? new System.Collections.Generic.List<string>();
            int colOffset = headers.FindIndex(h => string.Equals(h, header, StringComparison.OrdinalIgnoreCase));
            if (colOffset < 0) return false;
            int headerRows = (int)(table.HeaderRowCount?.Value ?? 1U);
            bool totals = table.TotalsRowShown?.Value ?? false;
            int totalsRows = totals ? 1 : 0;
            int startRow = r1 + headerRows;
            int endRow = r2 - totalsRows;
            int column = c1 + colOffset;
            var toSheet = destinationSheetForCellText ?? (text => text);
            for (int r = startRow; r <= endRow; r++)
            {
                if (!TryGetCellText(r, column, out var text) || string.IsNullOrWhiteSpace(text)) continue;
                string sheetName = toSheet(text);
                if (string.IsNullOrWhiteSpace(sheetName)) continue;
                string location = $"'{EscapeSheetNameForLink(sheetName)}'!{targetA1}";
                string disp = display?.Invoke(text) ?? text;
                SetInternalLink(r, column, location, disp, styled);
            }
            return true;
        }

        /// <summary>
        /// Links a column within a table to external URLs. The table range determines the row bounds.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="header">Header text of the column inside the table.</param>
        /// <param name="urlForCellText">Maps cell text to URL.</param>
        /// <param name="titleForCellText">Optional display selector; when null, a smart display (RFC/host) is used.</param>
        /// <param name="styled">Apply hyperlink styling (blue + underline).</param>
        public void LinkByHeaderToUrlsInTable(
            string tableName,
            string header,
            Func<string, string> urlForCellText,
            Func<string, string>? titleForCellText = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentNullException(nameof(tableName));
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (urlForCellText is null) throw new ArgumentNullException(nameof(urlForCellText));
            var tdp = _worksheetPart.TableDefinitionParts.FirstOrDefault(p => string.Equals(p.Table?.Name?.Value, tableName, StringComparison.OrdinalIgnoreCase));
            if (tdp?.Table == null) throw new InvalidOperationException($"Table '{tableName}' not found on sheet '{Name}'.");
            var table = tdp.Table;
            string? refA1 = table.Reference?.Value;
            if (string.IsNullOrWhiteSpace(refA1)) throw new InvalidOperationException($"Table '{tableName}' has no Reference.");
            var (r1, c1, r2, c2) = A1.ParseRange(refA1!);
            var headers = table.TableColumns?.Elements<DocumentFormat.OpenXml.Spreadsheet.TableColumn>().Select(tc => tc.Name?.Value ?? string.Empty).ToList() ?? new System.Collections.Generic.List<string>();
            int colOffset = headers.FindIndex(h => string.Equals(h, header, StringComparison.OrdinalIgnoreCase));
            if (colOffset < 0) throw new InvalidOperationException($"Header '{header}' not found in table '{tableName}'.");
            int headerRows = (int)(table.HeaderRowCount?.Value ?? 1U);
            bool totals = table.TotalsRowShown?.Value ?? false;
            int totalsRows = totals ? 1 : 0;
            int startRow = r1 + headerRows;
            int endRow = r2 - totalsRows;
            int column = c1 + colOffset;
            for (int r = startRow; r <= endRow; r++)
            {
                if (!TryGetCellText(r, column, out var text) || string.IsNullOrWhiteSpace(text)) continue;
                string url = urlForCellText(text);
                if (string.IsNullOrWhiteSpace(url)) continue;
                if (titleForCellText != null)
                    SetHyperlink(r, column, url, titleForCellText(text), styled);
                else
                    SetHyperlinkSmart(r, column, url, null, styled);
            }
        }

        /// <summary>
        /// Non-throwing variant of <see cref="LinkByHeaderToUrlsInTable"/>.
        /// Returns false when the table or header cannot be found.
        /// </summary>
        public bool TryLinkByHeaderToUrlsInTable(
            string tableName,
            string header,
            Func<string, string>? urlForCellText,
            Func<string, string>? titleForCellText = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(tableName) || string.IsNullOrWhiteSpace(header)) return false;
            if (urlForCellText is null) return false;
            var tdp = _worksheetPart.TableDefinitionParts.FirstOrDefault(p => string.Equals(p.Table?.Name?.Value, tableName, StringComparison.OrdinalIgnoreCase));
            var table = tdp?.Table; if (table == null) return false;
            string? refA1 = table.Reference?.Value; if (string.IsNullOrWhiteSpace(refA1)) return false;
            if (!A1.TryParseRange(refA1!, out var r1, out var c1, out var r2, out var c2)) return false;
            var headers = table.TableColumns?.Elements<DocumentFormat.OpenXml.Spreadsheet.TableColumn>().Select(tc => tc.Name?.Value ?? string.Empty).ToList() ?? new System.Collections.Generic.List<string>();
            int colOffset = headers.FindIndex(h => string.Equals(h, header, StringComparison.OrdinalIgnoreCase));
            if (colOffset < 0) return false;
            int headerRows = (int)(table.HeaderRowCount?.Value ?? 1U);
            bool totals = table.TotalsRowShown?.Value ?? false;
            int totalsRows = totals ? 1 : 0;
            int startRow = r1 + headerRows;
            int endRow = r2 - totalsRows;
            int column = c1 + colOffset;
            for (int r = startRow; r <= endRow; r++)
            {
                if (!TryGetCellText(r, column, out var text) || string.IsNullOrWhiteSpace(text)) continue;
                string url = urlForCellText(text);
                if (string.IsNullOrWhiteSpace(url)) continue;
                if (titleForCellText != null)
                    SetHyperlink(r, column, url, titleForCellText(text), styled);
                else
                    SetHyperlinkSmart(r, column, url, null, styled);
            }
            return true;
        }

        /// <summary>
        /// Links a column identified by <paramref name="header"/> within a rectangular A1 range to internal sheets.
        /// Uses the first row of the range as the header row and links rows r1+1..r2.
        /// </summary>
        /// <param name="rangeA1">A1 range (e.g., "A1:D50"). The first row is treated as the header row.</param>
        /// <param name="header">Header text to match (case-insensitive).</param>
        /// <param name="destinationSheetForCellText">Maps cell text to destination sheet name (defaults to identity).</param>
        /// <param name="targetA1">Destination cell on the target sheet (default "A1").</param>
        /// <param name="display">Optional display selector; defaults to the cell text.</param>
        /// <param name="styled">Apply hyperlink styling (blue + underline).</param>
        public void LinkByHeaderToInternalSheetsInRange(
            string rangeA1,
            string header,
            Func<string, string>? destinationSheetForCellText = null,
            string targetA1 = "A1",
            Func<string, string>? display = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(rangeA1)) throw new ArgumentNullException(nameof(rangeA1));
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            var (r1, c1, r2, c2) = A1.ParseRange(rangeA1);
            // Find header column within first row of range
            int headerCol = -1;
            for (int c = c1; c <= c2; c++)
            {
                if (TryGetCellText(r1, c, out var text) && !string.IsNullOrWhiteSpace(text) && string.Equals(text, header, StringComparison.OrdinalIgnoreCase))
                {
                    headerCol = c; break;
                }
            }
            if (headerCol < 0) throw new InvalidOperationException($"Header '{header}' not found in range '{rangeA1}'.");
            var toSheet = destinationSheetForCellText ?? (t => t);
            for (int r = r1 + 1; r <= r2; r++)
            {
                if (!TryGetCellText(r, headerCol, out var value) || string.IsNullOrWhiteSpace(value)) continue;
                string sheetName = toSheet(value);
                if (string.IsNullOrWhiteSpace(sheetName)) continue;
                string location = $"'{EscapeSheetNameForLink(sheetName)}'!{targetA1}";
                string disp = display?.Invoke(value) ?? value;
                SetInternalLink(r, headerCol, location, disp, styled);
            }
        }

        /// <summary>
        /// Links a column identified by <paramref name="header"/> within a rectangular A1 range to external URLs.
        /// Uses the first row of the range as the header row and links rows r1+1..r2.
        /// </summary>
        /// <param name="rangeA1">A1 range (e.g., "A1:D50"). The first row is treated as the header row.</param>
        /// <param name="header">Header text to match (case-insensitive).</param>
        /// <param name="urlForCellText">Maps cell text to URL.</param>
        /// <param name="titleForCellText">Optional display selector; when null, a smart display (RFC/host) is used.</param>
        /// <param name="styled">Apply hyperlink styling (blue + underline).</param>
        public void LinkByHeaderToUrlsInRange(
            string rangeA1,
            string header,
            Func<string, string> urlForCellText,
            Func<string, string>? titleForCellText = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(rangeA1)) throw new ArgumentNullException(nameof(rangeA1));
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (urlForCellText is null) throw new ArgumentNullException(nameof(urlForCellText));
            var (r1, c1, r2, c2) = A1.ParseRange(rangeA1);
            int headerCol = -1;
            for (int c = c1; c <= c2; c++)
            {
                if (TryGetCellText(r1, c, out var text) && !string.IsNullOrWhiteSpace(text) && string.Equals(text, header, StringComparison.OrdinalIgnoreCase))
                { headerCol = c; break; }
            }
            if (headerCol < 0) throw new InvalidOperationException($"Header '{header}' not found in range '{rangeA1}'.");
            for (int r = r1 + 1; r <= r2; r++)
            {
                if (!TryGetCellText(r, headerCol, out var value) || string.IsNullOrWhiteSpace(value)) continue;
                string url = urlForCellText(value);
                if (string.IsNullOrWhiteSpace(url)) continue;
                if (titleForCellText != null)
                    SetHyperlink(r, headerCol, url, titleForCellText(value), styled);
                else
                    SetHyperlinkSmart(r, headerCol, url, null, styled);
            }
        }

        /// <summary>
        /// Non-throwing variant of <see cref="LinkByHeaderToInternalSheetsInRange"/>.
        /// Returns false when the range cannot be parsed or the header is missing.
        /// </summary>
        public bool TryLinkByHeaderToInternalSheetsInRange(
            string rangeA1,
            string header,
            string targetA1 = "A1",
            Func<string, string>? destinationSheetForCellText = null,
            Func<string, string>? display = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(rangeA1) || string.IsNullOrWhiteSpace(header)) return false;
            if (!A1.TryParseRange(rangeA1, out var r1, out var c1, out var r2, out var c2)) return false;
            int headerCol = -1;
            for (int c = c1; c <= c2; c++)
            {
                if (TryGetCellText(r1, c, out var text) && !string.IsNullOrWhiteSpace(text) && string.Equals(text, header, StringComparison.OrdinalIgnoreCase))
                { headerCol = c; break; }
            }
            if (headerCol < 0) return false;
            var toSheet = destinationSheetForCellText ?? (t => t);
            for (int r = r1 + 1; r <= r2; r++)
            {
                if (!TryGetCellText(r, headerCol, out var value) || string.IsNullOrWhiteSpace(value)) continue;
                string sheetName = toSheet(value);
                if (string.IsNullOrWhiteSpace(sheetName)) continue;
                string location = $"'{EscapeSheetNameForLink(sheetName)}'!{targetA1}";
                string disp = display?.Invoke(value) ?? value;
                SetInternalLink(r, headerCol, location, disp, styled);
            }
            return true;
        }

        /// <summary>
        /// Non-throwing variant of <see cref="LinkByHeaderToUrlsInRange"/>.
        /// Returns false when the range cannot be parsed or the header is missing.
        /// </summary>
        public bool TryLinkByHeaderToUrlsInRange(
            string rangeA1,
            string header,
            Func<string, string>? urlForCellText,
            Func<string, string>? titleForCellText = null,
            bool styled = true)
        {
            if (string.IsNullOrWhiteSpace(rangeA1) || string.IsNullOrWhiteSpace(header)) return false;
            if (urlForCellText is null) return false;
            if (!A1.TryParseRange(rangeA1, out var r1, out var c1, out var r2, out var c2)) return false;
            int headerCol = -1;
            for (int c = c1; c <= c2; c++)
            {
                if (TryGetCellText(r1, c, out var text) && !string.IsNullOrWhiteSpace(text) && string.Equals(text, header, StringComparison.OrdinalIgnoreCase))
                { headerCol = c; break; }
            }
            if (headerCol < 0) return false;
            for (int r = r1 + 1; r <= r2; r++)
            {
                if (!TryGetCellText(r, headerCol, out var value) || string.IsNullOrWhiteSpace(value)) continue;
                string url = urlForCellText(value);
                if (string.IsNullOrWhiteSpace(url)) continue;
                if (titleForCellText != null)
                    SetHyperlink(r, headerCol, url, titleForCellText(value), styled);
                else
                    SetHyperlinkSmart(r, headerCol, url, null, styled);
            }
            return true;
        }
    }
}
