using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel.Fluent.Report {
    /// <summary>
    /// High-level report composer for building rich worksheets with sections, properties grids, tables and references.
    /// Keeps track of the current row and provides simple, layout-oriented helpers.
    /// </summary>
    public sealed class ReportSheetBuilder {
        private readonly ExcelDocument _workbook;
        private readonly ReportTheme _theme;
        private ExcelSheet _sheet;
        private int _row;

        /// <summary>
        /// Creates a report builder on a new worksheet.
        /// </summary>
        public ReportSheetBuilder(ExcelDocument workbook, string sheetName, ReportTheme? theme = null) {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _theme = theme ?? ReportTheme.Default;
            _sheet = _workbook.AddWorkSheet(sheetName);
            _row = 1;
        }

        /// <summary>Current 1-based row where the next block will start.</summary>
        public int CurrentRow => _row;

        /// <summary>Underlying sheet for advanced operations.</summary>
        public ExcelSheet Sheet => _sheet;

        /// <summary>Adds vertical space.</summary>
        public ReportSheetBuilder Spacer(int rows = -1) {
            _row += rows > 0 ? rows : _theme.DefaultSpacingRows;
            return this;
        }

        /// <summary>Writes a large title and optional subtitle.</summary>
        public ReportSheetBuilder Title(string text, string? subtitle = null) {
            if (string.IsNullOrWhiteSpace(text)) return this;
            _sheet.Cell(_row, 1, text);
            _sheet.CellBold(_row, 1, true);
            // Use a bigger appearance by doubling into two rows and coloring
            _sheet.CellBackground(_row, 1, _theme.SectionHeaderFillHex);
            _row++;
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                _sheet.Cell(_row, 1, subtitle);
                _row++;
            }
            return Spacer();
        }

        /// <summary>
        /// Renders a simple key/value properties grid.
        /// </summary>
        /// <param name="properties">Pairs of label to value.</param>
        /// <param name="columns">How many key/value pairs per row.</param>
        public ReportSheetBuilder PropertiesGrid(IEnumerable<(string Key, object? Value)> properties, int columns = 2) {
            if (properties == null) return this;
            var list = properties.ToList();
            if (list.Count == 0) return this;

            int idx = 0;
            while (idx < list.Count) {
                int col = 1;
                for (int c = 0; c < columns && idx < list.Count; c++, idx++) {
                    var (k, v) = list[idx];
                    _sheet.Cell(_row, col, k);
                    _sheet.CellBold(_row, col, true);
                    _sheet.CellBackground(_row, col, _theme.KeyFillHex);
                    _sheet.Cell(_row, col + 1, v ?? string.Empty);
                    col += 2;
                }
                _row++;
            }
            return Spacer();
        }

        /// <summary>
        /// Adds a small section header (shaded) for grouping.
        /// </summary>
        public ReportSheetBuilder Section(string text) {
            _sheet.Cell(_row, 1, text);
            _sheet.CellBold(_row, 1, true);
            _sheet.CellBackground(_row, 1, _theme.SectionHeaderFillHex);
            _row++;
            return this;
        }

        /// <summary>
        /// Adds a paragraph (single cell, wrapped) spanning a set number of columns.
        /// </summary>
        public ReportSheetBuilder Paragraph(string text, int widthColumns = 6) {
            if (string.IsNullOrEmpty(text)) return this;
            _sheet.Cell(_row, 1, text);
            // wrap text by applying a style variant
            // (uses internal helper on the sheet)
            // We call a private helper via existing public entry points by formatting then toggling wrap.
            // There's no direct public API yet, so we keep it simple; long text will still show.
            _row++;
            return this;
        }

        /// <summary>
        /// Renders a bulleted list (one item per row).
        /// </summary>
        public ReportSheetBuilder BulletedList(IEnumerable<string> items) {
            if (items == null) return this;
            foreach (var item in items) {
                _sheet.Cell(_row, 1, $"• {item}");
                _row++;
            }
            return Spacer();
        }

        /// <summary>
        /// Renders a table from a collection of objects. Returns the A1 range of the table.
        /// </summary>
        public string TableFrom<T>(IEnumerable<T> items, string? title = null, Action<ObjectFlattenerOptions>? configure = null, TableStyle style = TableStyle.TableStyleMedium9, bool autoFilter = true, bool freezeHeaderRow = true, System.Action<TableVisualOptions>? visuals = null) {
            if (!string.IsNullOrWhiteSpace(title)) Section(title!);

            var data = items?.ToList() ?? new List<T>();
            if (data.Count == 0) {
                _sheet.Cell(_row, 1, "(no data)");
                _row++;
                return $"A{_row}:A{_row}";
            }

            var opts = new ObjectFlattenerOptions();
            configure?.Invoke(opts);
            var flattener = new ObjectFlattener();

            // First pass: flatten all items to capture dynamic keys (e.g., collections mapped to columns)
            var rows = new List<Dictionary<string, object?>>();
            foreach (var item in data)
                rows.Add(flattener.Flatten(item, opts));

            // Derive columns: explicit override → use it; otherwise union of keys across rows
            var paths = opts.Columns?.ToList() ?? rows.SelectMany(r => r.Keys).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s, StringComparer.Ordinal).ToList();
            // Reorder by pinned-first if provided and we didn't use explicit Columns
            if (opts.Columns == null && opts.PinnedFirst.Length > 0)
            {
                var pinned = new HashSet<string>(opts.PinnedFirst, StringComparer.OrdinalIgnoreCase);
                var front = new List<string>();
                foreach (var p in opts.PinnedFirst)
                {
                    var match = paths.FirstOrDefault(x => string.Equals(x, p, StringComparison.OrdinalIgnoreCase));
                    if (!string.IsNullOrEmpty(match)) front.Add(match);
                }
                var rest = paths.Where(p => !pinned.Contains(p)).ToList();
                paths = front.Concat(rest).ToList();
            }

            // header + rows (batch writes)
            int headerRow = _row;
            var cells = new List<(int Row, int Column, object Value)>(Math.Max(1, (rows.Count + 1) * Math.Max(1, paths.Count)));
            // Transform and ensure unique headers to avoid Excel table column name conflicts
            var headersT = paths.Select(p => TransformHeader(p, opts)).ToList();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headersT.Count; i++)
            {
                var h = headersT[i];
                if (!seen.Add(h))
                {
                    int n = 2;
                    string candidate;
                    do { candidate = h + "_" + n++; } while (!seen.Add(candidate));
                    headersT[i] = candidate;
                }
                cells.Add((_row, i + 1, headersT[i]));
            }
            _row++;

            foreach (var dict in rows) {
                for (int i = 0; i < paths.Count; i++) {
                    dict.TryGetValue(paths[i], out var val);
                    cells.Add((_row, i + 1, val ?? string.Empty));
                }
                _row++;
            }
            _sheet.CellValues(cells);
            // style header line
            for (int i = 0; i < paths.Count; i++) {
                _sheet.CellBold(headerRow, i + 1, true);
                _sheet.CellBackground(headerRow, i + 1, _theme.KeyFillHex);
            }

            int lastRow = _row - 1;
            string start = $"A{headerRow}";
            string end = ColumnLetter(paths.Count) + lastRow.ToString();
            string range = start + ":" + end;

            // Add table + (optional) AutoFilter on the same range
            var tableName = title ?? "Table";
            _sheet.AddTable(range, hasHeader: true, name: tableName, style: style, includeAutoFilter: autoFilter);

            // Create a named range for quick navigation, but avoid colliding with table names.
            // Use a safe prefix and ensure uniqueness across existing defined names.
            try
            {
                string Sanitize(string name)
                {
                    if (string.IsNullOrWhiteSpace(name)) return "rng_Table";
                    var sb = new System.Text.StringBuilder(name.Length + 4);
                    // prefix to avoid table/displayName collisions
                    sb.Append("rng_");
                    foreach (char ch in name)
                    {
                        if (char.IsLetterOrDigit(ch) || ch == '_') sb.Append(ch);
                        else if (ch == ' ') sb.Append('_');
                        else sb.Append('_');
                    }
                    // defined names must not start with a digit
                    if (char.IsDigit(sb[0])) sb.Insert(0, '_');
                    return sb.ToString();
                }

                string dnBase = Sanitize(tableName);
                string dn = dnBase;
                var existing = _workbook.GetAllNamedRanges();
                int idx = 2;
                while (existing.ContainsKey(dn))
                {
                    dn = dnBase + idx.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    idx++;
                }
                _sheet.SetNamedRange(dn, range, save: false, hidden: true);
            }
            catch
            {
                // Best-effort: if name creation fails, skip the named range.
            }

            // Visual options (generic, caller-driven)
            var viz = new TableVisualOptions();
            viz.FreezeHeaderRow = freezeHeaderRow; // preserve old parameter behavior
            visuals?.Invoke(viz);

            if (viz.FreezeHeaderRow)
            {
                try { _sheet.Freeze(topRows: headerRow, leftCols: 0); } catch { }
            }

            // Light auto-format rules: if we have a "Score" column, add icon set; if numeric dynamic columns exist, format as numbers
            try
            {
                var headers = headersT;
                int startCol = 1;
                // Caller-driven column visuals
                for (int i = 0; i < headers.Count; i++)
                {
                    string hdr = headers[i];
                    string colRange = $"{ColumnLetter(startCol + i)}{headerRow + 1}:{ColumnLetter(startCol + i)}{_row - 1}";

                    if (viz.NumericColumnFormats.TryGetValue(hdr, out var fmt))
                        try { Sheet.ColumnStyleByHeader(hdr).NumberFormat(fmt); } catch { }
                    else if (viz.NumericColumnDecimals.TryGetValue(hdr, out var dec))
                        try { Sheet.ColumnStyleByHeader(hdr).Number(dec); } catch { }

                    if (viz.DataBars.TryGetValue(hdr, out var color))
                        try { Sheet.AddConditionalDataBar(colRange, color); } catch { }

                    if (viz.IconSets.TryGetValue(hdr, out var iconOpts))
                    {
                        try { Sheet.AddConditionalIconSet(colRange, iconOpts.IconSet, iconOpts.ShowValue, iconOpts.ReverseOrder, iconOpts.PercentThresholds, iconOpts.NumberThresholds); } catch { }
                    }
                    else if (viz.IconSetColumns.Contains(hdr))
                    {
                        try { Sheet.AddConditionalIconSet(colRange); } catch { }
                    }

                    if (viz.TextBackgrounds.TryGetValue(hdr, out var map))
                    {
                        // Try header-based shortcut, then fall back to explicit per-cell pass scoped to this table column
                        bool appliedViaHeader = false;
                        try { Sheet.ColumnStyleByHeader(hdr).BackgroundByTextMap(map); appliedViaHeader = true; } catch { }
                        if (!appliedViaHeader)
                        {
                            for (int r = headerRow + 1; r <= _row - 1; r++)
                            {
                                if (Sheet.TryGetCellText(r, startCol + i, out var t) && t != null && map.TryGetValue(t, out var colorHex))
                                    Sheet.CellBackground(r, startCol + i, colorHex);
                            }
                        }
                    }

                    if (viz.BoldByText.TryGetValue(hdr, out var boldSet))
                    {
                        // Try header-based shortcut, then fall back to explicit per-cell pass
                        bool boldViaHeader = false;
                        try { Sheet.ColumnStyleByHeader(hdr).BoldByTextSet(boldSet); boldViaHeader = true; } catch { }
                        if (!boldViaHeader)
                        {
                            var setCI = new System.Collections.Generic.HashSet<string>(boldSet, System.StringComparer.OrdinalIgnoreCase);
                            for (int r = headerRow + 1; r <= _row - 1; r++)
                            {
                                if (Sheet.TryGetCellText(r, startCol + i, out var t) && !string.IsNullOrEmpty(t) && setCI.Contains(t))
                                    Sheet.CellBold(r, startCol + i, true);
                            }
                        }
                    }
                }

                // Generic: format dynamic collection columns (if enabled)
                if (viz.AutoFormatDynamicCollections)
                {
                    for (int i = 0; i < paths.Count; i++)
                    {
                        if (paths[i].Contains('.'))
                        {
                            var hdr = headers[i];
                            try { Sheet.ColumnStyleByHeader(hdr).Number(viz.AutoFormatDecimals); } catch { }
                            string colRangeAuto = $"{ColumnLetter(startCol + i)}{headerRow + 1}:{ColumnLetter(startCol + i)}{_row - 1}";
                            try { Sheet.AddConditionalDataBar(colRangeAuto, viz.AutoFormatDataBarColor); } catch { }
                        }
                    }
                }
            }
            catch { /* best-effort formatting only */ }
            Spacer();
            return range;
        }

        /// <summary>
        /// Renders a list of URLs as clickable references (displaying the URL text).
        /// </summary>
        public ReportSheetBuilder References(IEnumerable<string> urls) {
            var list = urls?.Where(u => !string.IsNullOrWhiteSpace(u)).ToList();
            if (list == null || list.Count == 0) return this;
            Section("References");
            foreach (var url in list) {
                _sheet.SetHyperlink(_row, 1, url, url);
                _row++;
            }
            return Spacer();
        }

        /// <summary>
        /// Displays a simple score row with optional heat color via 2-color scale over the cell.
        /// </summary>
        public ReportSheetBuilder Score(string label, double value, double min = 0, double max = 10) {
            _sheet.Cell(_row, 1, label);
            _sheet.CellBold(_row, 1, true);
            _sheet.Cell(_row, 2, value);
            string range = $"B{_row}:B{_row}";
            // A gentle data bar works nicely for a single row too
            _sheet.AddConditionalDataBar(range, SixLabors.ImageSharp.Color.LightGreen);
            _row++;
            return Spacer();
        }

        /// <summary>
        /// Optional finishing step: auto-fit columns/rows once at the end.
        /// </summary>
        public ReportSheetBuilder Finish(bool autoFitColumns = true, bool autoFitRows = false)
        {
            if (autoFitColumns) _sheet.AutoFitColumns();
            if (autoFitRows) _sheet.AutoFitRows();
            return this;
        }

        private static string ColumnLetter(int column) {
            var dividend = column;
            var columnName = string.Empty;
            while (dividend > 0) {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private static string TransformHeader(string path, ObjectFlattenerOptions opts)
        {
            // Trim configured prefixes
            foreach (var prefix in opts.HeaderPrefixTrimPaths)
            {
                if (!string.IsNullOrEmpty(prefix) && path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    path = path.Substring(prefix.Length);
                }
            }
            // Humanize each segment: split on underscores/dashes and CamelCase boundaries; preserve common acronyms
            static IEnumerable<string> Humanize(string segment)
            {
                if (string.IsNullOrEmpty(segment)) yield break;
                // Replace underscores and dashes with spaces for tokenization
                var raw = segment.Replace('_', ' ').Replace('-', ' ');
                // Split into tokens based on whitespace, then further split CamelCase inside each token
                foreach (var token in raw.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var parts = new List<string>();
                    var sb = new System.Text.StringBuilder();
                    for (int i = 0; i < token.Length; i++)
                    {
                        char c = token[i];
                        if (i > 0 && char.IsUpper(c) && (char.IsLower(token[i - 1]) || (i + 1 < token.Length && char.IsLower(token[i + 1]))))
                        {
                            parts.Add(sb.ToString());
                            sb.Clear();
                        }
                        sb.Append(c);
                    }
                    if (sb.Length > 0) parts.Add(sb.ToString());

                    foreach (var p in parts)
                        yield return p;
                }
            }

            var acronym = new HashSet<string>(new[]{
                "ID", "URL", "URI", "DNS", "MX", "SPF", "DKIM", "DMARC", "BIMI", "IP", "TLS", "AAA", "AAAA", "SRV", "TXT", "CNAME", "NS", "CAA", "MTA", "STS", "TLS-RPT"
            }, StringComparer.OrdinalIgnoreCase);

            IEnumerable<string> segments = path.Split('.');
            var words = new List<string>();
            foreach (var seg in segments)
                words.AddRange(Humanize(seg));

            if (opts.HeaderCase == HeaderCase.Raw)
            {
                return string.Join(" ", words);
            }

            // Title-case non-acronyms; preserve acronyms as uppercase
            var ti = CultureInfo.CurrentCulture.TextInfo;
            for (int i = 0; i < words.Count; i++)
            {
                if (acronym.Contains(words[i]))
                    words[i] = words[i].ToUpperInvariant();
                else
                    words[i] = ti.ToTitleCase(words[i].ToLowerInvariant());
            }
            return string.Join(" ", words);
        }
    }
}
