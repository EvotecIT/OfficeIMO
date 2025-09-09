using System;
using System.Collections.Generic;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel.Fluent
{
    /// <summary>
    /// Neutral, Excel-first wrapper for building stacked worksheet content.
    /// </summary>
    public sealed class SheetComposer
    {
        private readonly ExcelDocument _workbook;
        private readonly SheetTheme _theme;
        private ExcelSheet _sheet;
        private int _row;

        /// <summary>
        /// Creates a new sheet and prepares a composer for adding content top-to-bottom.
        /// Also creates a hidden named range at the sheet top ("top_{sheet}") for navigation.
        /// </summary>
        /// <param name="workbook">Target workbook.</param>
        /// <param name="sheetName">Name of the sheet to create.</param>
        /// <param name="theme">Optional theme controlling colors and spacing.</param>
        public SheetComposer(ExcelDocument workbook, string sheetName, SheetTheme? theme = null)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _theme = theme ?? SheetTheme.Default;
            _sheet = _workbook.AddWorkSheet(sheetName);
            _row = 1;
            // Note: We no longer create hidden named ranges by default (to avoid Excel repairs on some versions).
            // Internal links use explicit "'Sheet'!A1" locations instead.
        }

        /// <summary>The underlying sheet created by this composer.</summary>
        public ExcelSheet Sheet => _sheet;
        /// <summary>Current row where the next write occurs (1-based).</summary>
        public int CurrentRow => _row;

        /// <summary>Adds vertical spacing by advancing the current row.</summary>
        /// <param name="rows">Number of rows to skip; when negative uses the theme default.</param>
        public SheetComposer Spacer(int rows = -1) { _row += rows > 0 ? rows : _theme.DefaultSpacingRows; return this; }

        /// <summary>Writes a title (bold with themed background) and optional subtitle.</summary>
        /// <param name="text">Title text.</param>
        /// <param name="subtitle">Optional subtitle on the next row.</param>
        public SheetComposer Title(string text, string? subtitle = null)
        {
            if (string.IsNullOrWhiteSpace(text)) return this;
            _sheet.Cell(_row, 1, text);
            _sheet.CellBold(_row, 1, true);
            _sheet.CellBackground(_row, 1, _theme.SectionHeaderFillHex);
            _row++;
            if (!string.IsNullOrWhiteSpace(subtitle))
            {
                _sheet.Cell(_row, 1, subtitle);
                _row++;
            }
            return Spacer();
        }

        /// <summary>Writes a section header (bold with themed background).</summary>
        public SheetComposer Section(string text)
        {
            _sheet.Cell(_row, 1, text);
            _sheet.CellBold(_row, 1, true);
            _sheet.CellBackground(_row, 1, _theme.SectionHeaderFillHex);
            _row++;
            return this;
        }

        /// <summary>Writes a paragraph-like line of text.</summary>
        /// <param name="text">Text to write.</param>
        /// <param name="widthColumns">Reserved for future wrapping/width behavior.</param>
        public SheetComposer Paragraph(string text, int widthColumns = 6)
        {
            if (string.IsNullOrEmpty(text)) return this;
            _sheet.Cell(_row, 1, text);
            _row++;
            return this;
        }

        /// <summary>
        /// Inserts a simple callout (admonition) band consisting of a bold title row and a body row,
        /// shaded according to the <paramref name="kind"/>. Does not merge cells; applies background across
        /// the first <paramref name="widthColumns"/> columns for a banner effect.
        /// Supported kinds: info, success, warning, error/critical.
        /// </summary>
        public SheetComposer Callout(string kind, string title, string body, int widthColumns = 8)
        {
            string fill = kind?.Trim().ToLowerInvariant() switch
            {
                "success" => "#D4EDDA",
                "warning" => "#FFF3CD",
                "error" => "#F8D7DA",
                "critical" => "#F8D7DA",
                _ => "#E8F4FF" // info/default
            };

            if (!string.IsNullOrWhiteSpace(title))
            {
                _sheet.Cell(_row, 1, title);
                _sheet.CellBold(_row, 1, true);
                for (int c = 1; c <= Math.Max(1, widthColumns); c++) _sheet.CellBackground(_row, c, fill);
                _row++;
            }

            if (!string.IsNullOrWhiteSpace(body))
            {
                // Encourage wrapping by injecting a soft break if the line is long and has no breaks
                string text = body;
                if (!text.Contains("\n") && text.Length > 120)
                {
                    int cut = 120;
                    text = text.Insert(cut, "\n");
                }
                _sheet.Cell(_row, 1, text);
                for (int c = 1; c <= Math.Max(1, widthColumns); c++) _sheet.CellBackground(_row, c, fill);
                _row++;
            }
            return Spacer();
        }

        /// <summary>
        /// Renders a key/value list in a compact two-column grid. Alias for <see cref="PropertiesGrid"/>.
        /// </summary>
        /// <summary>Alias for <see cref="PropertiesGrid"/>.</summary>
        public SheetComposer DefinitionList(IEnumerable<(string Key, object? Value)> items, int columns = 2)
            => PropertiesGrid(items, columns);

        /// <summary>Renders a compact grid of key/value pairs (two columns per entry).</summary>
        /// <param name="properties">Sequence of (Key, Value) items.</param>
        /// <param name="columns">How many key/value pairs to render per row.</param>
        public SheetComposer PropertiesGrid(IEnumerable<(string Key, object? Value)> properties, int columns = 2)
        {
            if (properties == null) return this;
            var list = new List<(string Key, object? Value)>(properties);
            if (list.Count == 0) return this;
            int idx = 0;
            while (idx < list.Count)
            {
                int col = 1;
                for (int c = 0; c < columns && idx < list.Count; c++, idx++)
                {
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

        /// <summary>Writes a simple bulleted list, one item per row.</summary>
        public SheetComposer BulletedList(IEnumerable<string> items)
        {
            if (items == null) return this;
            foreach (var item in items)
            {
                _sheet.Cell(_row, 1, $"• {item}");
                _row++;
            }
            return Spacer();
        }

        /// <summary>
        /// Flattens a sequence of objects into a table and renders it with a header row.
        /// Returns the A1 range used for the table.
        /// </summary>
        public string TableFrom<T>(IEnumerable<T> items, string? title = null,
            Action<ObjectFlattenerOptions>? configure = null,
            TableStyle style = TableStyle.TableStyleMedium9,
            bool autoFilter = true,
            bool freezeHeaderRow = true,
            Action<TableVisualOptions>? visuals = null)
        {
            if (!string.IsNullOrWhiteSpace(title)) Section(title!);

            var data = items?.ToList() ?? new List<T>();
            if (data.Count == 0)
            {
                _sheet.Cell(_row, 1, "(no data)");
                _row++;
                return $"A{_row}:A{_row}";
            }

            var opts = new ObjectFlattenerOptions();
            configure?.Invoke(opts);
            var flattener = new ObjectFlattener();

            var rows = new List<Dictionary<string, object?>>();
            foreach (var item in data)
                rows.Add(flattener.Flatten(item, opts));

            var paths = opts.Columns?.ToList() ?? rows.SelectMany(r => r.Keys).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s, StringComparer.Ordinal).ToList();
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

            int headerRow = _row;
            var cells = new List<(int Row, int Column, object Value)>(Math.Max(1, (rows.Count + 1) * Math.Max(1, paths.Count)));
            var headersT = paths.Select(p => TransformHeader(p, opts)).ToList();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headersT.Count; i++)
            {
                var h = headersT[i];
                if (!seen.Add(h))
                {
                    int n = 2; string candidate;
                    do { candidate = h + "_" + n++; } while (!seen.Add(candidate));
                    headersT[i] = candidate;
                }
                cells.Add((_row, i + 1, headersT[i]));
            }
            _row++;

            foreach (var dict in rows)
            {
                for (int i = 0; i < paths.Count; i++)
                {
                    dict.TryGetValue(paths[i], out var val);
                    cells.Add((_row, i + 1, val ?? string.Empty));
                }
                _row++;
            }
            _sheet.CellValues(cells);
            for (int i = 0; i < paths.Count; i++)
            {
                _sheet.CellBold(headerRow, i + 1, true);
                _sheet.CellBackground(headerRow, i + 1, _theme.KeyFillHex);
            }

            int lastRow = _row - 1;
            string start = $"A{headerRow}";
            string end = ColumnLetter(paths.Count) + lastRow.ToString();
            string range = start + ":" + end;

            var tableName = title ?? "Table";
            _sheet.AddTable(range, hasHeader: true, name: tableName, style: style, includeAutoFilter: autoFilter);

            // Avoid creating a named range for the table by default; callers can add one explicitly if needed.

            var viz = new TableVisualOptions();
            viz.FreezeHeaderRow = freezeHeaderRow; visuals?.Invoke(viz);
            if (viz.FreezeHeaderRow) { try { _sheet.Freeze(topRows: headerRow, leftCols: 0); } catch { } }

            try
            {
                var headers = headersT; int startCol = 1;
                for (int i = 0; i < headers.Count; i++)
                {
                    string hdr = headers[i];
                    string colRange = $"{ColumnLetter(startCol + i)}{headerRow + 1}:{ColumnLetter(startCol + i)}{_row - 1}";

                    if (viz.NumericColumnFormats.TryGetValue(hdr, out var fmt))
                    {
                        if (_sheet.TryGetColumnIndexByHeader(hdr, out _))
                            _sheet.ColumnStyleByHeader(hdr).NumberFormat(fmt);
                    }
                    else if (viz.NumericColumnDecimals.TryGetValue(hdr, out var dec))
                    {
                        if (_sheet.TryGetColumnIndexByHeader(hdr, out _))
                            _sheet.ColumnStyleByHeader(hdr).Number(dec);
                    }

                    if (viz.DataBars.TryGetValue(hdr, out var color))
                        try { _sheet.AddConditionalDataBar(colRange, color); } catch { }

                    if (viz.IconSets.TryGetValue(hdr, out var iconOpts))
                        try { _sheet.AddConditionalIconSet(colRange, iconOpts.IconSet, iconOpts.ShowValue, iconOpts.ReverseOrder, iconOpts.PercentThresholds, iconOpts.NumberThresholds); } catch { }
                    else if (viz.IconSetColumns.Contains(hdr))
                        try { _sheet.AddConditionalIconSet(colRange); } catch { }

                    if (viz.TextBackgrounds.TryGetValue(hdr, out var map))
                    {
                        if (_sheet.TryGetColumnIndexByHeader(hdr, out _))
                        {
                            _sheet.ColumnStyleByHeader(hdr).BackgroundByTextMap(map);
                        }
                        else
                        {
                            for (int r = headerRow + 1; r <= _row - 1; r++)
                                if (_sheet.TryGetCellText(r, startCol + i, out var t) && t != null && map.TryGetValue(t, out var colorHex))
                                    _sheet.CellBackground(r, startCol + i, colorHex);
                        }
                    }
                    if (viz.BoldByText.TryGetValue(hdr, out var boldSet))
                    {
                        if (_sheet.TryGetColumnIndexByHeader(hdr, out _))
                        {
                            _sheet.ColumnStyleByHeader(hdr).BoldByTextSet(boldSet);
                        }
                        else
                        {
                            var setCI = new HashSet<string>(boldSet, StringComparer.OrdinalIgnoreCase);
                            for (int r = headerRow + 1; r <= _row - 1; r++)
                                if (_sheet.TryGetCellText(r, startCol + i, out var t) && !string.IsNullOrEmpty(t) && setCI.Contains(t))
                                    _sheet.CellBold(r, startCol + i, true);
                        }
                    }
                }

                if (viz.AutoFormatDynamicCollections)
                {
                    for (int i = 0; i < paths.Count; i++)
                    {
                        if (paths[i].Contains('.'))
                        {
                            var hdr = headers[i];
                            if (_sheet.TryGetColumnIndexByHeader(hdr, out _))
                                _sheet.ColumnStyleByHeader(hdr).Number(viz.AutoFormatDecimals);
                            string colRangeAuto = $"{ColumnLetter(startCol + i)}{headerRow + 1}:{ColumnLetter(startCol + i)}{_row - 1}";
                            try { _sheet.AddConditionalDataBar(colRangeAuto, viz.AutoFormatDataBarColor); } catch { }
                        }
                    }
                }
            }
            catch { }
            Spacer();
            return range;
        }

        /// <summary>Writes a simple References section with each URL as a hyperlink.</summary>
        public SheetComposer References(IEnumerable<string> urls)
        {
            var list = urls?.Where(u => !string.IsNullOrWhiteSpace(u)).ToList();
            if (list != null && list.Count > 0)
            {
                Section("References");
                foreach (var url in list) { _sheet.SetHyperlinkSmart(_row, 1, url); _row++; }
                Spacer();
            }
            return this;
        }

        /// <summary>Writes a labeled numeric score with a data bar visualization.</summary>
        public SheetComposer Score(string label, double value, double min = 0, double max = 10)
        {
            _sheet.Cell(_row, 1, label);
            _sheet.CellBold(_row, 1, true);
            _sheet.Cell(_row, 2, value);
            string range = $"B{_row}:B{_row}";
            _sheet.AddConditionalDataBar(range, SixLabors.ImageSharp.Color.LightGreen);
            _row++;
            return Spacer();
        }

        /// <summary>Applies optional autofit operations and returns the composer.</summary>
        public SheetComposer Finish(bool autoFitColumns = true, bool autoFitRows = false)
        {
            if (autoFitColumns) _sheet.AutoFitColumns();
            if (autoFitRows) _sheet.AutoFitRows();
            return this;
        }

        /// <summary>
        /// Inserts a section header with an optional anchor (named range) and a back-to-top internal link.
        /// </summary>
        /// <param name="text">Section header text.</param>
        /// <param name="anchorName">Optional anchor name; generated from text when null.</param>
        /// <param name="backToTopLink">Whether to add a back-to-top link under the section.</param>
        /// <param name="backToTopText">Link text for the back-to-top link.</param>
        public SheetComposer SectionWithAnchor(string text, string? anchorName = null, bool backToTopLink = true, string backToTopText = "↑ Top")
        {
            Section(text);
            // Named range anchor intentionally omitted (explicit cell links are used instead).
            if (backToTopLink)
            {
                try {
                    string topName = SanitizeName($"top_{_sheet.Name}");
                    // Internal link: '#SheetName!A1' also works, but using named range for clarity
                    _sheet.SetInternalLink(_row, 1, $"'{_sheet.Name}'!A1", backToTopText);
                    _row++;
                } catch { }
            }
            return this;
        }

        /// <summary>Configures header/footer content and images via a builder.</summary>
        /// <param name="configure">Callback that applies header/footer settings.</param>
        public SheetComposer HeaderFooter(Action<HeaderFooterBuilder> configure)
        {
            if (configure == null) return this;
            var b = new HeaderFooterBuilder();
            configure(b);
            b.Apply(_sheet);
            return this;
        }

        private static string ColumnLetter(int column)
        {
            var dividend = column; var columnName = string.Empty;
            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private static string TransformHeader(string path, ObjectFlattenerOptions opts)
        {
            foreach (var prefix in opts.HeaderPrefixTrimPaths)
                if (!string.IsNullOrEmpty(prefix) && path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    path = path.Substring(prefix.Length);
            static IEnumerable<string> Humanize(string segment)
            {
                if (string.IsNullOrEmpty(segment)) yield break;
                var raw = segment.Replace('_', ' ').Replace('-', ' ');
                foreach (var token in raw.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var parts = new List<string>(); var sb = new System.Text.StringBuilder();
                    for (int i = 0; i < token.Length; i++)
                    {
                        char c = token[i];
                        if (i > 0 && char.IsUpper(c) && (char.IsLower(token[i - 1]) || (i + 1 < token.Length && char.IsLower(token[i + 1]))))
                        { parts.Add(sb.ToString()); sb.Clear(); }
                        sb.Append(c);
                    }
                    if (sb.Length > 0) parts.Add(sb.ToString());
                    foreach (var p in parts) yield return p;
                }
            }
            var acronym = new HashSet<string>(new[]{
                "ID","URL","URI","DNS","MX","SPF","DKIM","DMARC","BIMI","IP","TLS","AAA","AAAA","SRV","TXT","CNAME","NS","CAA","MTA","STS","TLS-RPT"
            }, StringComparer.OrdinalIgnoreCase);
            IEnumerable<string> segments = path.Split('.'); var words = new List<string>();
            foreach (var seg in segments) words.AddRange(Humanize(seg));
            var ti = System.Globalization.CultureInfo.CurrentCulture.TextInfo;
            for (int i = 0; i < words.Count; i++)
                words[i] = acronym.Contains(words[i]) ? words[i].ToUpperInvariant() : ti.ToTitleCase(words[i].ToLowerInvariant());
            return string.Join(" ", words);
        }

        private static string SanitizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "_";
            var sb = new System.Text.StringBuilder(name.Length + 8);
            foreach (char ch in name)
            {
                if (char.IsLetterOrDigit(ch) || ch == '_' || ch == '.') sb.Append(ch);
                else if (char.IsWhiteSpace(ch) || ch == '-' || ch == '/') sb.Append('_');
            }
            var s = sb.ToString();
            if (string.IsNullOrEmpty(s)) s = "_";
            if (char.IsDigit(s[0])) s = "_" + s;
            return s.Length > 255 ? s.Substring(0, 255) : s;
        }
    }
}
