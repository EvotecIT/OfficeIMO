using System.Text.RegularExpressions;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Parses an A1 range and returns 1-based bounds (r1, c1, r2, c2).
        /// </summary>
        /// <param name="a1">A1 range without a sheet prefix (e.g., "A2:D20").</param>
        /// <returns>Tuple of (r1, c1, r2, c2) with normalized bounds.</returns>
        /// <exception cref="ArgumentException">Thrown when the input is not a valid A1 range.</exception>
        public (int r1, int c1, int r2, int c2) GetRangeBounds(string a1) {
            return ExcelA1.ParseRange(a1);
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
        public void ForEachRow(string a1, Action<int> action) {
            var (r1, _, r2, _) = ExcelA1.ParseRange(a1);
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
        public void ForEachColumn(string a1, Action<int> action) {
            var (_, c1, _, c2) = ExcelA1.ParseRange(a1);
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
        public void LinkCellsToInternalSheets(string a1, Func<string, string> destinationSheetForCellText, string targetA1 = "A1", Func<string, string>? display = null, bool styled = true) {
            ArgumentNullException.ThrowIfNull(destinationSheetForCellText);
            var (r1, c1, r2, c2) = ExcelA1.ParseRange(a1);
            for (int r = r1; r <= r2; r++) {
                for (int c = c1; c <= c2; c++) {
                    if (!TryGetCellText(r, c, out var text) || string.IsNullOrWhiteSpace(text)) continue;
                    string sheetName = destinationSheetForCellText(text);
                    if (string.IsNullOrWhiteSpace(sheetName)) continue;
                    string location = ResolveInternalSheetLinkLocation(sheetName, targetA1);
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
        public void SetHyperlinkSmart(int row, int column, string url, string? title = null, bool style = true) {
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
        public void SetHyperlinkHost(int row, int column, string url, bool style = true) {
            string display = TryGetHost(url, out var host) ? host! : url;
            SetHyperlink(row, column, url, display, style);
        }
        /// <example>
        /// sheet.SetHyperlinkHost(7, 1, "https://learn.microsoft.com/office/open-xml/"); // displays "learn.microsoft.com"
        /// </example>

        private static string GuessLinkDisplay(string url) {
            // RFC pattern: https://datatracker.ietf.org/doc/html/rfc7208
            var r = Regex.Match(url ?? string.Empty, @"rfc(\d{3,5})", RegexOptions.IgnoreCase);
            if (r.Success) return "RFC " + r.Groups[1].Value;
            if (TryGetHost(url, out var host)) return host!;
            return url ?? string.Empty;
        }

        private static bool TryGetHost(string? url, out string? host) {
            host = null;
            if (string.IsNullOrWhiteSpace(url)) return false;
            if (Uri.TryCreate(url, UriKind.Absolute, out var uri)) {
                host = uri.Host;
                return !string.IsNullOrEmpty(host);
            }
            return false;
        }

        private string ResolveInternalSheetLinkLocation(string requestedSheetName, string targetA1) {
            return SheetNameLookup.BuildInternalLocation(_excelDocument.Sheets, requestedSheetName, targetA1);
        }

        /// <summary>
        /// Turns each non-empty cell under a header into a link to a worksheet or URL built from the cell text.
        /// </summary>
        /// <param name="header">Header text of the column to link (case-insensitive).</param>
        /// <param name="target">Where each cell points: <see cref="ExcelLinkTarget.Sheet"/> or <see cref="ExcelLinkTarget.Url"/>.</param>
        /// <param name="options">Optional range or table scope, row bounds, display text and styling.</param>
        /// <returns>The linked column and count, the skipped cells, or why nothing was linked.</returns>
        /// <example>
        /// sheet.LinkByHeader("Domain", ExcelLinkTarget.Sheet());
        /// sheet.LinkByHeader("RFC", ExcelLinkTarget.Url(rfc => $"https://datatracker.ietf.org/doc/html/{rfc}"),
        ///     new ExcelLinkByHeaderOptions { Table = "Summary" });
        /// </example>
        public ExcelLinkByHeaderResult LinkByHeader(string header, ExcelLinkTarget target, ExcelLinkByHeaderOptions? options = null) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            ArgumentNullException.ThrowIfNull(target);
            var opts = options ?? new ExcelLinkByHeaderOptions();
            if (!string.IsNullOrWhiteSpace(opts.Range) && !string.IsNullOrWhiteSpace(opts.Table)) {
                throw new ArgumentException("Set either Range or Table, not both.", nameof(options));
            }

            int column;
            int firstRow;
            int lastRow;
            if (!string.IsNullOrWhiteSpace(opts.Table)) {
                string tableName = opts.Table!;
                var table = _worksheetPart.TableDefinitionParts
                    .Select(part => part.Table)
                    .FirstOrDefault(candidate => string.Equals(candidate?.Name?.Value, tableName, StringComparison.OrdinalIgnoreCase));
                if (table == null || !ExcelA1.TryParseRange(table.Reference?.Value ?? string.Empty, out int r1, out int c1, out int r2, out _)) {
                    return ExcelLinkByHeaderResult.Failed($"Table '{tableName}' was not found on sheet '{Name}'.");
                }

                var headers = table.TableColumns?.Elements<DocumentFormat.OpenXml.Spreadsheet.TableColumn>()
                    .Select(tableColumn => tableColumn.Name?.Value ?? string.Empty).ToList() ?? new List<string>();
                int offset = headers.FindIndex(name => string.Equals(name, header, StringComparison.OrdinalIgnoreCase));
                if (offset < 0) return ExcelLinkByHeaderResult.Failed($"Header '{header}' was not found in table '{tableName}'.");
                column = c1 + offset;
                firstRow = r1 + (int)(table.HeaderRowCount?.Value ?? 1U);
                lastRow = r2 - (table.TotalsRowShown?.Value == true ? 1 : 0);
            } else if (!string.IsNullOrWhiteSpace(opts.Range)) {
                if (!ExcelA1.TryParseRange(opts.Range!, out int r1, out int c1, out int r2, out int c2)) {
                    throw new ArgumentException($"Range '{opts.Range}' is not a valid A1 range.", nameof(options));
                }

                column = -1;
                for (int c = c1; c <= c2; c++) {
                    if (TryGetCellText(r1, c, out var text) && !string.IsNullOrWhiteSpace(text) && string.Equals(text, header, StringComparison.OrdinalIgnoreCase)) {
                        column = c;
                        break;
                    }
                }

                if (column < 0) return ExcelLinkByHeaderResult.Failed($"Header '{header}' was not found in range '{opts.Range}'.");
                firstRow = r1 + 1;
                lastRow = r2;
            } else {
                if (!TryGetColumnIndexByHeader(header, out column)) {
                    return ExcelLinkByHeaderResult.Failed($"Header '{header}' was not found on sheet '{Name}'.");
                }

                firstRow = opts.FirstRow;
                if (opts.LastRow is int requestedLastRow && requestedLastRow > 0) {
                    lastRow = requestedLastRow;
                } else {
                    var (usedFirstRow, _, usedLastRow, _) = ExcelA1.ParseRange(GetUsedRangeA1());
                    lastRow = usedLastRow;
                    if (firstRow < usedFirstRow) firstRow = usedFirstRow + 1;
                }
            }

            int linked = 0;
            List<string>? skipped = null;
            Func<string, string> map = target.Map ?? (text => text);
            for (int row = firstRow; row <= lastRow; row++) {
                if (!TryGetCellText(row, column, out var text) || string.IsNullOrWhiteSpace(text)) continue;
                string destination = map(text);
                if (string.IsNullOrWhiteSpace(destination)) {
                    (skipped ??= new List<string>()).Add(ExcelA1.CellReference(row, column));
                    continue;
                }

                if (target.IsUrl) {
                    if (opts.Display != null) SetHyperlink(row, column, destination, opts.Display(text), opts.Styled);
                    else SetHyperlinkSmart(row, column, destination, null, opts.Styled);
                } else {
                    string location = ResolveInternalSheetLinkLocation(destination, target.TargetCell);
                    SetInternalLink(row, column, location, opts.Display?.Invoke(text) ?? text, opts.Styled);
                }

                linked++;
            }

            return new ExcelLinkByHeaderResult(null, column, linked, (IReadOnlyList<string>?)skipped ?? Array.Empty<string>());
        }
    }
}
