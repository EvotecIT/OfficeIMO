using System.Text;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

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

    internal string GetHeaderFooterSectionText(ExcelHeaderFooterKind kind, ExcelHeaderFooterPosition position) =>
        ReadWorksheetObject(worksheet => {
            OfficeXmlNode value = worksheet.Child(WorksheetMarkup, "headerFooter");
            var sections = ParseHeaderFooterSections(GetHeaderFooterText(value, kind) ?? string.Empty);
            return position == ExcelHeaderFooterPosition.Left ? sections.L
                : position == ExcelHeaderFooterPosition.Center ? sections.C
                : sections.R;
        });

    internal ExcelHeaderFooterImageInfo? GetHeaderFooterSectionImage(ExcelHeaderFooterKind kind, ExcelHeaderFooterPosition position) {
        if (kind != ExcelHeaderFooterKind.OddHeader && kind != ExcelHeaderFooterKind.OddFooter) return null;
        string shapeId = (position == ExcelHeaderFooterPosition.Left ? "L" : position == ExcelHeaderFooterPosition.Center ? "C" : "R")
            + (kind == ExcelHeaderFooterKind.OddHeader ? "H" : "F");
        return ReadHeaderFooterImages(ExcelImageExportOptions.DefaultMaximumTotalSourceImageBytes).Images
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

        EditHeaderFooter(value => {
            var current = ParseHeaderFooterSections(GetHeaderFooterText(value, kind) ?? string.Empty);
            string? rebuilt = BuildHeaderFooterSections(
                position == ExcelHeaderFooterPosition.Left ? text : current.L,
                position == ExcelHeaderFooterPosition.Center ? text : current.C,
                position == ExcelHeaderFooterPosition.Right ? text : current.R);
            SetHeaderFooterText(value, kind, rebuilt);
            if (text.Length > 0) {
                value.SetBooleanAttribute(kind is ExcelHeaderFooterKind.FirstHeader or ExcelHeaderFooterKind.FirstFooter
                    ? "differentFirst" : "differentOddEven", true);
            }
        });
    }

    private static string HeaderFooterElementName(ExcelHeaderFooterKind kind) => kind switch {
        ExcelHeaderFooterKind.OddHeader => "oddHeader",
        ExcelHeaderFooterKind.OddFooter => "oddFooter",
        ExcelHeaderFooterKind.FirstHeader => "firstHeader",
        ExcelHeaderFooterKind.FirstFooter => "firstFooter",
        ExcelHeaderFooterKind.EvenHeader => "evenHeader",
        _ => "evenFooter"
    };

    private static readonly OfficeXmlChildOrder HeaderFooterOrder = new(
        new[] { WorksheetMarkup, WorksheetMarkup, WorksheetMarkup, WorksheetMarkup, WorksheetMarkup, WorksheetMarkup },
        new[] { "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter" },
        new[] { 0, 1, 2, 3, 4, 5 });

    private static string? GetHeaderFooterText(OfficeXmlNode value, ExcelHeaderFooterKind kind) =>
        value.ChildText(WorksheetMarkup, HeaderFooterElementName(kind));

    private static void SetHeaderFooterText(OfficeXmlNode value, ExcelHeaderFooterKind kind, string? text) =>
        value.SetChildText(WorksheetMarkup, HeaderFooterElementName(kind), text, HeaderFooterOrder);

    private void UpdateHeaderFooterSections(bool header, string? left, string? center, string? right,
        bool? differentFirstPage, bool? differentOddEven, bool? alignWithMargins, bool? scaleWithDocument) {
        if (left == null && center == null && right == null && differentFirstPage == null && differentOddEven == null
            && alignWithMargins == null && scaleWithDocument == null) return;
        EditHeaderFooter(value => {
            bool enableFirst = differentFirstPage == true && value.BooleanAttribute("differentFirst") != true;
            bool enableOddEven = differentOddEven == true && value.BooleanAttribute("differentOddEven") != true;
            if (left != null || center != null || right != null) {
                ExcelHeaderFooterKind kind = header ? ExcelHeaderFooterKind.OddHeader : ExcelHeaderFooterKind.OddFooter;
                var current = ParseHeaderFooterSections(GetHeaderFooterText(value, kind) ?? string.Empty);
                SetHeaderFooterText(value, kind, BuildHeaderFooterSections(left ?? current.L, center ?? current.C, right ?? current.R));
            }
            if (differentFirstPage.HasValue) value.SetBooleanAttribute("differentFirst", differentFirstPage);
            if (differentOddEven.HasValue) value.SetBooleanAttribute("differentOddEven", differentOddEven);
            if (alignWithMargins.HasValue) value.SetBooleanAttribute("alignWithMargins", alignWithMargins);
            if (scaleWithDocument.HasValue) value.SetBooleanAttribute("scaleWithDoc", scaleWithDocument);
            if (enableFirst) {
                InheritHeaderFooterText(value, ExcelHeaderFooterKind.FirstHeader, ExcelHeaderFooterKind.OddHeader);
                InheritHeaderFooterText(value, ExcelHeaderFooterKind.FirstFooter, ExcelHeaderFooterKind.OddFooter);
            }
            if (enableOddEven) {
                InheritHeaderFooterText(value, ExcelHeaderFooterKind.EvenHeader, ExcelHeaderFooterKind.OddHeader);
                InheritHeaderFooterText(value, ExcelHeaderFooterKind.EvenFooter, ExcelHeaderFooterKind.OddFooter);
            }
        });
    }

    private static void InheritHeaderFooterText(OfficeXmlNode value, ExcelHeaderFooterKind destination, ExcelHeaderFooterKind source) {
        if (GetHeaderFooterText(value, destination) == null && GetHeaderFooterText(value, source) is string text) {
            SetHeaderFooterText(value, destination, text);
        }
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
        // Excel treats text before the first section marker as centered content.
        char section = 'C';
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
