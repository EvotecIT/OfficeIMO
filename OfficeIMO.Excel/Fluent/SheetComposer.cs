using System;
using System.Collections.Generic;
using OfficeIMO.Excel.Utilities;

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
            _sheet = _workbook.AddWorkSheet(sheetName);
            _row = 1;
            // Note: We no longer create hidden named ranges by default (to avoid Excel repairs on some versions).
            // Internal links use explicit "'Sheet'!A1" locations instead.
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
    }
}
