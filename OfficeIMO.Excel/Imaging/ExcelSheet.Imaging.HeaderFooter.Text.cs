using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryResolveHeaderFooterText(string? text, int pageNumber, int pageCount, DateTime headerFooterDateTime, out HeaderFooterTextSection normalized) {
            normalized = HeaderFooterTextSection.Empty;
            if (string.IsNullOrWhiteSpace(text)) {
                return true;
            }

            var runs = new List<HeaderFooterTextRun>();
            var builder = new StringBuilder(text!.Length);
            bool bold = false;
            bool italic = false;
            bool underline = false;
            bool hasFormatting = false;
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch != '&') {
                    builder.Append(ch);
                    continue;
                }

                if (i + 1 >= text.Length) {
                    return false;
                }

                char token = text[++i];
                if (token == '&') {
                    builder.Append('&');
                } else if (token == 'B') {
                    FlushHeaderFooterTextRun(runs, builder, bold, italic, underline);
                    bold = !bold;
                    hasFormatting = true;
                } else if (token == 'I') {
                    FlushHeaderFooterTextRun(runs, builder, bold, italic, underline);
                    italic = !italic;
                    hasFormatting = true;
                } else if (token == 'U') {
                    FlushHeaderFooterTextRun(runs, builder, bold, italic, underline);
                    underline = !underline;
                    hasFormatting = true;
                } else if (token == 'P') {
                    builder.Append(pageNumber.ToString(CultureInfo.InvariantCulture));
                } else if (token == 'N') {
                    builder.Append(pageCount.ToString(CultureInfo.InvariantCulture));
                } else if (token == 'D') {
                    builder.Append(FormatHeaderFooterDate(headerFooterDateTime));
                } else if (token == 'T') {
                    builder.Append(FormatHeaderFooterTime(headerFooterDateTime));
                } else if (token == 'A') {
                    builder.Append(Name);
                } else if (token == 'F') {
                    if (!TryGetWorkbookFileName(out string fileName)) {
                        return false;
                    }

                    builder.Append(fileName);
                } else if (token == 'Z') {
                    if (!TryGetWorkbookPathPrefix(out string pathPrefix)) {
                        return false;
                    }

                    builder.Append(pathPrefix);
                } else if (token == '[') {
                    int end = text.IndexOf(']', i + 1);
                    if (end < 0) {
                        return false;
                    }

                    string fieldName = text.Substring(i + 1, end - i - 1);
                    if (!TryAppendHeaderFooterField(builder, fieldName, pageNumber, pageCount, headerFooterDateTime)) {
                        return false;
                    }

                    i = end;
                } else {
                    return false;
                }
            }

            FlushHeaderFooterTextRun(runs, builder, bold, italic, underline);
            normalized = HeaderFooterTextSection.Create(runs, hasFormatting);
            return true;
        }

        private static void FlushHeaderFooterTextRun(List<HeaderFooterTextRun> runs, StringBuilder builder, bool bold, bool italic, bool underline) {
            if (builder.Length == 0) {
                return;
            }

            runs.Add(new HeaderFooterTextRun(builder.ToString(), bold, italic, underline));
            builder.Clear();
        }

        private bool TryAppendHeaderFooterField(StringBuilder builder, string fieldName, int pageNumber, int pageCount, DateTime headerFooterDateTime) {
            if (string.Equals(fieldName, "Page", StringComparison.OrdinalIgnoreCase)) {
                builder.Append(pageNumber.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (string.Equals(fieldName, "Pages", StringComparison.OrdinalIgnoreCase)) {
                builder.Append(pageCount.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (string.Equals(fieldName, "Tab", StringComparison.OrdinalIgnoreCase)) {
                builder.Append(Name);
                return true;
            }

            if (string.Equals(fieldName, "Date", StringComparison.OrdinalIgnoreCase)) {
                builder.Append(FormatHeaderFooterDate(headerFooterDateTime));
                return true;
            }

            if (string.Equals(fieldName, "Time", StringComparison.OrdinalIgnoreCase)) {
                builder.Append(FormatHeaderFooterTime(headerFooterDateTime));
                return true;
            }

            if (string.Equals(fieldName, "File", StringComparison.OrdinalIgnoreCase)) {
                if (!TryGetWorkbookFileName(out string fileName)) {
                    return false;
                }

                builder.Append(fileName);
                return true;
            }

            if (string.Equals(fieldName, "Path", StringComparison.OrdinalIgnoreCase)) {
                if (!TryGetWorkbookPathPrefix(out string pathPrefix)) {
                    return false;
                }

                builder.Append(pathPrefix);
                return true;
            }

            return false;
        }

        private static string FormatHeaderFooterDate(DateTime headerFooterDateTime) =>
            headerFooterDateTime.ToString("d", CultureInfo.CurrentCulture);

        private static string FormatHeaderFooterTime(DateTime headerFooterDateTime) =>
            headerFooterDateTime.ToString("t", CultureInfo.CurrentCulture);

        private bool TryGetWorkbookFileName(out string fileName) {
            fileName = string.Empty;
            string path = _excelDocument.FilePath;
            if (string.IsNullOrWhiteSpace(path)) {
                return false;
            }

            fileName = Path.GetFileName(path);
            return !string.IsNullOrWhiteSpace(fileName);
        }

        private bool TryGetWorkbookPathPrefix(out string pathPrefix) {
            pathPrefix = string.Empty;
            string path = _excelDocument.FilePath;
            if (string.IsNullOrWhiteSpace(path)) {
                return false;
            }

            string? directory = Path.GetDirectoryName(Path.GetFullPath(path));
            if (string.IsNullOrWhiteSpace(directory)) {
                return true;
            }

            pathPrefix = directory!.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ||
                directory.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                ? directory
                : directory + Path.DirectorySeparatorChar;
            return true;
        }

        private sealed class HeaderFooterTextSection {
            internal static readonly HeaderFooterTextSection Empty = new HeaderFooterTextSection(Array.Empty<HeaderFooterTextRun>(), string.Empty, hasFormatting: false);

            private HeaderFooterTextSection(IReadOnlyList<HeaderFooterTextRun> runs, string text, bool hasFormatting) {
                Runs = runs;
                Text = text;
                HasFormatting = hasFormatting;
            }

            internal IReadOnlyList<HeaderFooterTextRun> Runs { get; }
            internal string Text { get; }
            internal bool HasFormatting { get; }
            internal bool HasText => !string.IsNullOrWhiteSpace(Text);

            internal static HeaderFooterTextSection Create(IReadOnlyList<HeaderFooterTextRun> runs, bool hasFormatting) {
                if (runs.Count == 0) {
                    return Empty;
                }

                string text = string.Concat(runs.Select(run => run.Text)).Trim();
                if (string.IsNullOrWhiteSpace(text)) {
                    return Empty;
                }

                var normalizedRuns = new List<HeaderFooterTextRun>(runs.Count);
                for (int index = 0; index < runs.Count; index++) {
                    HeaderFooterTextRun run = runs[index];
                    string runText = run.Text;
                    if (index == 0) {
                        runText = runText.TrimStart();
                    }

                    if (index == runs.Count - 1) {
                        runText = runText.TrimEnd();
                    }

                    if (runText.Length > 0) {
                        normalizedRuns.Add(new HeaderFooterTextRun(runText, run.Bold, run.Italic, run.Underline));
                    }
                }

                return normalizedRuns.Count == 0
                    ? Empty
                    : new HeaderFooterTextSection(normalizedRuns.AsReadOnly(), text, hasFormatting);
            }

            internal IReadOnlyList<OfficeRichTextRun> ToOfficeRuns(double fontSize, OfficeColor color) {
                var runs = new List<OfficeRichTextRun>(Runs.Count);
                for (int index = 0; index < Runs.Count; index++) {
                    HeaderFooterTextRun run = Runs[index];
                    runs.Add(new OfficeRichTextRun(run.Text, fontSize, color, run.Bold, run.Italic, run.Underline));
                }

                return runs.AsReadOnly();
            }
        }

        private readonly struct HeaderFooterTextRun {
            internal HeaderFooterTextRun(string text, bool bold, bool italic, bool underline) {
                Text = text;
                Bold = bold;
                Italic = italic;
                Underline = underline;
            }

            internal string Text { get; }
            internal bool Bold { get; }
            internal bool Italic { get; }
            internal bool Underline { get; }
        }
    }
}
