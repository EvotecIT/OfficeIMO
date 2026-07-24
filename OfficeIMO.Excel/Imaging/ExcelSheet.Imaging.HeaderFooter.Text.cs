using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int MaximumHeaderFooterFontFamilyCharacters = 256;

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
            bool strikethrough = false;
            OfficeColor? color = null;
            double? fontSize = null;
            string? fontFamily = null;
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
                    if (i + 1 < text.Length && char.IsDigit(text[i + 1])) {
                        if (!TryReadHeaderFooterFontSizeToken(text, i + 1, out double parsedFontSize, out int tokenEnd)) {
                            return false;
                        }

                        FlushHeaderFooterTextRun(runs, builder, bold, italic, underline, strikethrough, color, fontSize, fontFamily);
                        fontSize = parsedFontSize;
                        i = tokenEnd;
                        hasFormatting = true;
                    } else {
                        builder.Append('&');
                    }
                } else if (token == 'B') {
                    FlushHeaderFooterTextRun(runs, builder, bold, italic, underline, strikethrough, color, fontSize, fontFamily);
                    bold = !bold;
                    hasFormatting = true;
                } else if (token == 'I') {
                    FlushHeaderFooterTextRun(runs, builder, bold, italic, underline, strikethrough, color, fontSize, fontFamily);
                    italic = !italic;
                    hasFormatting = true;
                } else if (token == 'U') {
                    FlushHeaderFooterTextRun(runs, builder, bold, italic, underline, strikethrough, color, fontSize, fontFamily);
                    underline = !underline;
                    hasFormatting = true;
                } else if (token == 'S') {
                    FlushHeaderFooterTextRun(runs, builder, bold, italic, underline, strikethrough, color, fontSize, fontFamily);
                    strikethrough = !strikethrough;
                    hasFormatting = true;
                } else if (token == 'K') {
                    if (!TryReadHeaderFooterColorToken(text, i + 1, out OfficeColor parsedColor)) {
                        return false;
                    }

                    FlushHeaderFooterTextRun(runs, builder, bold, italic, underline, strikethrough, color, fontSize, fontFamily);
                    color = parsedColor;
                    i += 6;
                    hasFormatting = true;
                } else if (char.IsDigit(token)) {
                    if (!TryReadHeaderFooterFontSizeToken(text, i, out double parsedFontSize, out int tokenEnd)) {
                        return false;
                    }

                    FlushHeaderFooterTextRun(runs, builder, bold, italic, underline, strikethrough, color, fontSize, fontFamily);
                    fontSize = parsedFontSize;
                    i = tokenEnd;
                    hasFormatting = true;
                } else if (token == '"') {
                    if (!TryReadHeaderFooterFontFamilyToken(text, i, out string parsedFontFamily, out bool? parsedBold, out bool? parsedItalic, out int tokenEnd)) {
                        return false;
                    }

                    FlushHeaderFooterTextRun(runs, builder, bold, italic, underline, strikethrough, color, fontSize, fontFamily);
                    fontFamily = parsedFontFamily;
                    if (parsedBold.HasValue) {
                        bold = parsedBold.Value;
                    }

                    if (parsedItalic.HasValue) {
                        italic = parsedItalic.Value;
                    }

                    i = tokenEnd;
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
                } else if (token == 'G') {
                    FlushHeaderFooterTextRun(runs, builder, bold, italic, underline, strikethrough, color, fontSize, fontFamily);
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

            FlushHeaderFooterTextRun(runs, builder, bold, italic, underline, strikethrough, color, fontSize, fontFamily);
            normalized = HeaderFooterTextSection.Create(runs, hasFormatting);
            return true;
        }

        private static void FlushHeaderFooterTextRun(List<HeaderFooterTextRun> runs, StringBuilder builder, bool bold, bool italic, bool underline, bool strikethrough, OfficeColor? color, double? fontSize, string? fontFamily) {
            if (builder.Length == 0) {
                return;
            }

            runs.Add(new HeaderFooterTextRun(builder.ToString(), bold, italic, underline, strikethrough, color, fontSize, fontFamily));
            builder.Clear();
        }

        private static bool TryReadHeaderFooterColorToken(string text, int start, out OfficeColor color) {
            color = default;
            if (start + 6 > text.Length) {
                return false;
            }

            string hex = text.Substring(start, 6);
            return OfficeColor.TryParseHex(hex, out color);
        }

        private static bool TryReadHeaderFooterFontSizeToken(string text, int start, out double fontSize, out int tokenEnd) {
            fontSize = 0D;
            tokenEnd = start;
            while (tokenEnd + 1 < text.Length && char.IsDigit(text[tokenEnd + 1])) {
                tokenEnd++;
            }

            string value = text.Substring(start, tokenEnd - start + 1);
            return double.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out fontSize) && fontSize > 0D;
        }

        private static bool TryReadHeaderFooterFontFamilyToken(string text, int openingQuoteIndex, out string fontFamily, out bool? bold, out bool? italic, out int tokenEnd) {
            fontFamily = string.Empty;
            bold = null;
            italic = null;
            tokenEnd = openingQuoteIndex;

            int end = text.IndexOf('"', openingQuoteIndex + 1);
            if (end < 0) {
                return false;
            }

            if (end - openingQuoteIndex - 1 > MaximumHeaderFooterFontFamilyCharacters) {
                return false;
            }

            string token = text.Substring(openingQuoteIndex + 1, end - openingQuoteIndex - 1).Trim();
            if (token.Length == 0) {
                return false;
            }

            string style = string.Empty;
            int comma = token.IndexOf(',');
            if (comma >= 0) {
                style = token.Substring(comma + 1).Trim();
                token = token.Substring(0, comma).Trim();
            }

            if (token.Length == 0 || !TryResolveHeaderFooterFontStyle(style, out bold, out italic)) {
                return false;
            }

            fontFamily = token;
            tokenEnd = end;
            return true;
        }

        private static bool TryResolveHeaderFooterFontStyle(string style, out bool? bold, out bool? italic) {
            bold = null;
            italic = null;
            if (string.IsNullOrWhiteSpace(style)) {
                return true;
            }

            string normalized = style.Trim().Replace("-", string.Empty).Replace(" ", string.Empty);
            if (string.Equals(normalized, "Regular", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "Normal", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "Standard", StringComparison.OrdinalIgnoreCase)) {
                bold = false;
                italic = false;
                return true;
            }

            if (string.Equals(normalized, "Bold", StringComparison.OrdinalIgnoreCase)) {
                bold = true;
                return true;
            }

            if (string.Equals(normalized, "Italic", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "Oblique", StringComparison.OrdinalIgnoreCase)) {
                italic = true;
                return true;
            }

            if (string.Equals(normalized, "BoldItalic", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "ItalicBold", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "BoldOblique", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "ObliqueBold", StringComparison.OrdinalIgnoreCase)) {
                bold = true;
                italic = true;
                return true;
            }

            return false;
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
            string? path = _excelDocument.FilePath;
            if (string.IsNullOrWhiteSpace(path)) {
                return false;
            }

            fileName = Path.GetFileName(path);
            return !string.IsNullOrWhiteSpace(fileName);
        }

        private bool TryGetWorkbookPathPrefix(out string pathPrefix) {
            pathPrefix = string.Empty;
            string? path = _excelDocument.FilePath;
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
                        normalizedRuns.Add(new HeaderFooterTextRun(runText, run.Bold, run.Italic, run.Underline, run.Strikethrough, run.Color, run.FontSize, run.FontFamily));
                    }
                }

                return normalizedRuns.Count == 0
                    ? Empty
                    : new HeaderFooterTextSection(normalizedRuns.AsReadOnly(), text, hasFormatting);
            }

            internal IReadOnlyList<OfficeRichTextRun> ToOfficeRuns(double fontSize, OfficeColor color, string defaultFontFamily) {
                var runs = new List<OfficeRichTextRun>(Runs.Count);
                for (int index = 0; index < Runs.Count; index++) {
                    HeaderFooterTextRun run = Runs[index];
                    runs.Add(new OfficeRichTextRun(
                        run.Text,
                        ResolveHeaderFooterRunFontSize(fontSize, run.FontSize),
                        run.Color ?? color,
                        run.Bold,
                        run.Italic,
                        run.Underline,
                        fontFamily: string.IsNullOrWhiteSpace(run.FontFamily) ? defaultFontFamily : run.FontFamily,
                        strikethrough: run.Strikethrough));
                }

                return runs.AsReadOnly();
            }

            internal double GetMaxResolvedFontSize(double defaultFontSize) {
                double max = defaultFontSize;
                for (int index = 0; index < Runs.Count; index++) {
                    max = Math.Max(max, ResolveHeaderFooterRunFontSize(defaultFontSize, Runs[index].FontSize));
                }

                return max;
            }

            private static double ResolveHeaderFooterRunFontSize(double defaultFontSize, double? requestedFontSize) {
                if (!requestedFontSize.HasValue) {
                    return defaultFontSize;
                }

                double scale = defaultFontSize / HeaderFooterFontSize;
                return Math.Max(1D, requestedFontSize.Value * scale);
            }
        }

        private readonly struct HeaderFooterTextRun {
            internal HeaderFooterTextRun(string text, bool bold, bool italic, bool underline, bool strikethrough, OfficeColor? color, double? fontSize, string? fontFamily) {
                Text = text;
                Bold = bold;
                Italic = italic;
                Underline = underline;
                Strikethrough = strikethrough;
                Color = color;
                FontSize = fontSize;
                FontFamily = fontFamily;
            }

            internal string Text { get; }
            internal bool Bold { get; }
            internal bool Italic { get; }
            internal bool Underline { get; }
            internal bool Strikethrough { get; }
            internal OfficeColor? Color { get; }
            internal double? FontSize { get; }
            internal string? FontFamily { get; }
        }
    }
}
