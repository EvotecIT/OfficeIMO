using System;
using System.Collections.Generic;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel.Fluent
{
    /// <summary>
    /// Neutral, Excel-first wrapper for building stacked worksheet content.
    /// </summary>
    public sealed partial class SheetComposer
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
            // Use sanitized sheet names to avoid Excel repair prompts for invalid characters/length
            _sheet = _workbook.AddWorkSheet(sheetName, SheetNameValidationMode.Sanitize);
            _row = 1;
            // Create a sheet-local top anchor so callers/tests can rely on a defined name at A1.
            // Keep it simple and safe: local to this sheet, absolute A1.
            // Use a simple A1 reference so the normalizer can expand to $A$1
            try { _workbook.SetNamedRange($"top_{SanitizeName(_sheet.Name)}", "A1", _sheet, save: true, hidden: true); } catch { }
        }

        /// <summary>The underlying sheet created by this composer.</summary>
        public ExcelSheet Sheet => _sheet;
        /// <summary>Current row where the next write occurs (1-based).</summary>
        public int CurrentRow => _row;


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

        /// <summary>
        /// Applies clamped widths and optional wrap based on table headers for a given A1 range.
        /// The first row of the range is treated as the header row.
        /// </summary>
        /// <param name="a1Range">A1 range returned from TableFrom (e.g., "A10:F42").</param>
        /// <param name="configure">Configure sizing options and header categories.</param>
        public SheetComposer ApplyColumnSizing(string? a1Range, Action<ColumnSizingOptions> configure)
        {
            if (string.IsNullOrWhiteSpace(a1Range)) return this;
            var (fromCol, fromRow) = OfficeIMO.Excel.A1.ParseCellRef(a1Range!.Split(':')[0]);
            var (toCol, toRow) = OfficeIMO.Excel.A1.ParseCellRef(a1Range!.Split(':')[1]);
            if (fromCol <= 0 || fromRow <= 0 || toCol <= 0 || toRow <= 0) return this;

            // Read header texts from first row
            var headers = new Dictionary<int, string>(toCol - fromCol + 1);
            for (int c = fromCol; c <= toCol; c++)
            {
                string text = string.Empty;
                try { Sheet.TryGetCellText(fromRow, c, out text); } catch { }
                headers[c] = (text ?? string.Empty).Trim();
            }

            var opts = new ColumnSizingOptions();
            configure?.Invoke(opts);

            bool IsMatch(HashSet<string> set, string header)
            {
                if (string.IsNullOrEmpty(header)) return false;
                if (set.Contains(header)) return true;
                // Loose match: allow simple aliases like "TLS 1.3" → "TLS13"
                var norm = header.Replace(" ", string.Empty).Replace("-", string.Empty);
                foreach (var s in set)
                {
                    var sn = s.Replace(" ", string.Empty).Replace("-", string.Empty);
                    if (norm.Equals(sn, StringComparison.OrdinalIgnoreCase)) return true;
                }
                return false;
            }

            var autoFitTargets = new HashSet<int>();

            foreach (var kv in headers)
            {
                int col = kv.Key; string h = kv.Value;
                bool hasHeader = !string.IsNullOrEmpty(h);

                double? width = null;
                if (hasHeader && opts.WidthByHeader.TryGetValue(h, out var explicitWidth))
                {
                    width = explicitWidth;
                }
                else if (hasHeader && IsMatch(opts.ShortHeaders, h)) width = opts.ShortWidth;
                else if (hasHeader && IsMatch(opts.NumericHeaders, h)) width = opts.NumericWidth;
                else if (hasHeader && IsMatch(opts.LongHeaders, h)) width = opts.LongWidth;

                bool shouldWrap = hasHeader && (IsMatch(opts.WrapHeaders, h) || IsMatch(opts.LongHeaders, h));
                if (shouldWrap && width == null)
                {
                    width = opts.WrapWidth;
                }

                if (width == null && !opts.AutoFitRemainingColumns)
                {
                    width = opts.MediumWidth;
                }

                if (width.HasValue)
                {
                    try { Sheet.SetColumnWidth(col, width.Value); } catch { }
                }

                if (shouldWrap)
                {
                    try
                    {
                        if (width.HasValue) Sheet.WrapCells(fromRow + 1, toRow, col, width.Value);
                        else Sheet.WrapCells(fromRow + 1, toRow, col);
                    }
                    catch { }
                }

                bool explicitAutoFit = hasHeader && IsMatch(opts.AutoFitHeaders, h);
                if (!shouldWrap && !width.HasValue && (explicitAutoFit || opts.AutoFitRemainingColumns))
                {
                    autoFitTargets.Add(col);
                }
            }

            if (autoFitTargets.Count > 0)
            {
                try { Sheet.AutoFitColumnsFor(autoFitTargets); } catch { }
            }
            return this;
        }
    }
}
