using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Updates running-header sections and flags, retaining omitted sections, the footer and authored page variants.</summary>
    /// <para>Null leaves a section or setting unchanged; empty text clears a section. Different-page flags apply to
    /// both header and footer. Changing a variant flag from false to true fills missing content from odd-page text; existing variant
    /// text remains intact. First-page and even-page sections are edited through SetHeaderFooterSectionText.</para>
    internal void SetHeaderCore(string? left = null, string? center = null, string? right = null,
        bool? differentFirstPage = null, bool? differentOddEven = null, bool? alignWithMargins = null,
        bool? scaleWithDocument = null) => UpdateHeaderFooterSections(true, left, center, right,
            differentFirstPage, differentOddEven, alignWithMargins, scaleWithDocument);

    /// <summary>Updates running-footer sections and flags, retaining omitted sections, the header and authored page variants.</summary>
    /// <para>Null leaves a section or setting unchanged; empty text clears a section. Different-page flags apply to
    /// both header and footer. Existing first/even-page content is retained, including when a variant is disabled.</para>
    internal void SetFooterCore(string? left = null, string? center = null, string? right = null,
        bool? differentFirstPage = null, bool? differentOddEven = null, bool? alignWithMargins = null,
        bool? scaleWithDocument = null) => UpdateHeaderFooterSections(false, left, center, right,
            differentFirstPage, differentOddEven, alignWithMargins, scaleWithDocument);

    internal string GetHeaderFooterSectionText(ExcelHeaderFooterKind kind, ExcelHeaderFooterPosition position) {
        HeaderFooter? value = WorksheetRoot.GetFirstChild<HeaderFooter>();
        var sections = ParseHeaderFooterSections(GetHeaderFooterText(value, kind) ?? string.Empty);
        return position == ExcelHeaderFooterPosition.Left ? sections.L
            : position == ExcelHeaderFooterPosition.Center ? sections.C
            : sections.R;
    }

    internal ExcelHeaderFooterImageInfo? GetHeaderFooterSectionImage(ExcelHeaderFooterKind kind, ExcelHeaderFooterPosition position) {
        if (kind != ExcelHeaderFooterKind.OddHeader && kind != ExcelHeaderFooterKind.OddFooter) return null;
        string shapeId = (position == ExcelHeaderFooterPosition.Left ? "L" : position == ExcelHeaderFooterPosition.Center ? "C" : "R")
            + (kind == ExcelHeaderFooterKind.OddHeader ? "H" : "F");
        return ReadHeaderFooterImages(ExcelImageExportOptions.DefaultMaximumTotalSourceImageBytes)
            .TryGetValue(shapeId, out ExcelHeaderFooterImageInfo? image) ? image : null;
    }

    /// <summary>Replaces one section; first-page and even-page text turns its variant on, as Excel's dialog does.</summary>
    internal void SetHeaderFooterSectionText(ExcelHeaderFooterKind kind, ExcelHeaderFooterPosition position, string text) {
        if (kind == ExcelHeaderFooterKind.OddHeader || kind == ExcelHeaderFooterKind.OddFooter) {
            UpdateHeaderFooterSections(kind == ExcelHeaderFooterKind.OddHeader,
                position == ExcelHeaderFooterPosition.Left ? text : null,
                position == ExcelHeaderFooterPosition.Center ? text : null,
                position == ExcelHeaderFooterPosition.Right ? text : null,
                null, null, null, null);
            return;
        }

        WriteLock(() => {
            HeaderFooter value = EnsureHeaderFooter();
            var current = ParseHeaderFooterSections(GetHeaderFooterText(value, kind) ?? string.Empty);
            string? rebuilt = BuildHeaderFooterSections(
                position == ExcelHeaderFooterPosition.Left ? text : current.L,
                position == ExcelHeaderFooterPosition.Center ? text : current.C,
                position == ExcelHeaderFooterPosition.Right ? text : current.R);
            switch (kind) {
                case ExcelHeaderFooterKind.FirstHeader: value.FirstHeader = rebuilt == null ? null : new FirstHeader(rebuilt); break;
                case ExcelHeaderFooterKind.FirstFooter: value.FirstFooter = rebuilt == null ? null : new FirstFooter(rebuilt); break;
                case ExcelHeaderFooterKind.EvenHeader: value.EvenHeader = rebuilt == null ? null : new EvenHeader(rebuilt); break;
                default: value.EvenFooter = rebuilt == null ? null : new EvenFooter(rebuilt); break;
            }

            if (text.Length > 0) {
                if (kind == ExcelHeaderFooterKind.FirstHeader || kind == ExcelHeaderFooterKind.FirstFooter) value.DifferentFirst = true;
                else value.DifferentOddEven = true;
            }

            CleanupHeaderFooterPictureArtifacts();
        });
    }

    private static string? GetHeaderFooterText(HeaderFooter? value, ExcelHeaderFooterKind kind) => kind switch {
        ExcelHeaderFooterKind.OddHeader => value?.OddHeader?.Text,
        ExcelHeaderFooterKind.OddFooter => value?.OddFooter?.Text,
        ExcelHeaderFooterKind.FirstHeader => value?.FirstHeader?.Text,
        ExcelHeaderFooterKind.FirstFooter => value?.FirstFooter?.Text,
        ExcelHeaderFooterKind.EvenHeader => value?.EvenHeader?.Text,
        _ => value?.EvenFooter?.Text
    };

    private void UpdateHeaderFooterSections(bool header, string? left, string? center, string? right,
        bool? differentFirstPage, bool? differentOddEven, bool? alignWithMargins, bool? scaleWithDocument) {
        if (left == null && center == null && right == null && differentFirstPage == null && differentOddEven == null
            && alignWithMargins == null && scaleWithDocument == null) return;
        WriteLock(() => {
            HeaderFooter value = EnsureHeaderFooter();
            bool enableFirst = differentFirstPage == true && value.DifferentFirst?.Value != true;
            bool enableOddEven = differentOddEven == true && value.DifferentOddEven?.Value != true;
            if (left != null || center != null || right != null) {
                var current = ParseHeaderFooterSections(header ? value.OddHeader?.Text ?? string.Empty : value.OddFooter?.Text ?? string.Empty);
                string? text = BuildHeaderFooterSections(left ?? current.L, center ?? current.C, right ?? current.R);
                if (header) value.OddHeader = text == null ? null : new OddHeader(text);
                else value.OddFooter = text == null ? null : new OddFooter(text);
            }
            if (differentFirstPage.HasValue) value.DifferentFirst = differentFirstPage.Value;
            if (differentOddEven.HasValue) value.DifferentOddEven = differentOddEven.Value;
            if (alignWithMargins.HasValue) value.AlignWithMargins = alignWithMargins.Value;
            if (scaleWithDocument.HasValue) value.ScaleWithDoc = scaleWithDocument.Value;
            // Missing variants inherit the current running text; authored variants are never replaced implicitly.
            if (enableFirst) {
                if (value.FirstHeader == null && value.OddHeader != null) value.FirstHeader = new FirstHeader(value.OddHeader.Text);
                if (value.FirstFooter == null && value.OddFooter != null) value.FirstFooter = new FirstFooter(value.OddFooter.Text);
            }
            if (enableOddEven) {
                if (value.EvenHeader == null && value.OddHeader != null) value.EvenHeader = new EvenHeader(value.OddHeader.Text);
                if (value.EvenFooter == null && value.OddFooter != null) value.EvenFooter = new EvenFooter(value.OddFooter.Text);
            }
            CleanupHeaderFooterPictureArtifacts();
        });
    }

    private static IEnumerable<(int Offset, char Code)> HeaderFooterCodes(string text) {
        for (int index = 0; index + 1 < text.Length; index++) {
            if (text[index] != '&') continue;
            int start = index;
            char code = text[++index];
            if (code == '&') continue;
            if (code == '"') {
                int end = text.IndexOf('"', index + 1);
                index = end < 0 ? text.Length : end;
                continue;
            }
            yield return (start, code);
        }
    }

    private static bool HasHeaderFooterCode(string? text, char code) {
        if (string.IsNullOrEmpty(text)) return false;
        foreach (var token in HeaderFooterCodes(text!)) if (token.Code == code) return true;
        return false;
    }

    private static (string L, string C, string R) ParseHeaderFooterSections(string text) {
        var left = new StringBuilder();
        var center = new StringBuilder();
        var right = new StringBuilder();
        char section = '\0';
        int start = 0;
        void Append(int end) {
            StringBuilder? target = section == 'L' ? left : section == 'C' ? center : section == 'R' ? right : null;
            target?.Append(text, start, end - start);
        }
        foreach (var token in HeaderFooterCodes(text)) {
            if (token.Code != 'L' && token.Code != 'C' && token.Code != 'R') continue;
            Append(token.Offset);
            section = token.Code;
            start = token.Offset + 2;
        }
        Append(text.Length);
        return (left.ToString(), center.ToString(), right.ToString());
    }

    private static string? BuildHeaderFooterSections(string? left, string? center, string? right) {
        var builder = new StringBuilder();
        if (!string.IsNullOrEmpty(left)) builder.Append("&L").Append(EscapeHeaderFooter(left));
        if (!string.IsNullOrEmpty(center)) builder.Append("&C").Append(EscapeHeaderFooter(center));
        if (!string.IsNullOrEmpty(right)) builder.Append("&R").Append(EscapeHeaderFooter(right));
        return builder.Length == 0 ? null : builder.ToString();
    }

    private static string EscapeHeaderFooter(string? text) {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        bool IsTokenStarter(char c) {
            // Recognize common Excel header/footer tokens following '&' to avoid escaping them
            if (c >= '0' && c <= '9') {
                return true;
            }

            switch (c) {
                case 'L':
                case 'C':
                case 'R': // section markers
                case 'P':
                case 'N':
                case 'D':
                case 'T':
                case 'A':
                case 'F':
                case 'Z': // page, pages, date, time, sheet, file, path
                case 'G': // picture placeholder
                case 'K': // color: &Krrggbb
                case 'B':
                case 'I':
                case 'E':
                case 'U':
                case 'S': // bold, italic, underline, strike
                case 'X':
                case 'Y': // superscript, subscript
                case '[': // bracketed fields such as &[Page], &[Pages], &[Tab]
                    return true;
            }
            return false;
        }

        var t = text!;
        var sb = new StringBuilder(t.Length + 8);
        for (int i = 0; i < t.Length; i++) {
            char ch = t[i];
            if (ch == '&') {
                if (i + 1 < t.Length) {
                    char n = t[i + 1];
                    if (n == '&') { sb.Append("&&"); i++; continue; }
                    if (n == '"') {
                        // A quoted font descriptor is one code, including any literal ampersands in its name.
                        int end = t.IndexOf('"', i + 2);
                        int length = (end < 0 ? t.Length : end + 1) - i;
                        sb.Append(t, i, length);
                        i += length - 1;
                        continue;
                    }
                    if (IsTokenStarter(n)) { sb.Append('&'); continue; }
                }
                // literal &
                sb.Append("&&");
            } else sb.Append(ch);
        }
        return sb.ToString();
    }

}
