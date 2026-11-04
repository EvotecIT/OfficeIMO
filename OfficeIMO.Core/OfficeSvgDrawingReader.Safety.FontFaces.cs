using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    /// <summary>Distinguishes bounded embedded-font descriptors from declarations that change painted element geometry.</summary>
    private static bool ContainsStylesheetRasterGeometryDeclaration(string? stylesheet,
        SvgEmbeddedFontImports fontImports, ref int declarationCount, ref int ruleCount) {
        if (string.IsNullOrWhiteSpace(stylesheet)) return false;
        if (HasSvgCssCommentInsideUnquotedUrl(stylesheet!)) return true;
        string css = StripCssComments(stylesheet!, out bool complete);
        if (!complete || ContainsNonSvgCssWhitespace(css)) return true;

        int cursor = 0;
        while (cursor < css.Length) {
            int open = FindSvgCssCharacter(css, '{', cursor);
            if (open < 0) return !string.IsNullOrWhiteSpace(css.Substring(cursor));
            int close = FindSvgCssBlockEnd(css, open + 1);
            if (close < 0) return true;
            string selector = css.Substring(cursor, open - cursor).Trim();
            string body = css.Substring(open + 1, close - open - 1);
            if (!IsBalancedRasterCssBlock(selector) || !IsBalancedRasterCssBlock(body)) return true;
            IReadOnlyList<SvgCssDeclaration> declarations = ParseSvgCssDeclarations(body, ref declarationCount, out complete);
            if (!complete || SplitSvgCssTopLevel(body, ';').Count(item => !string.IsNullOrWhiteSpace(item)) != declarations.Count) return true;

            if (selector.Equals("@font-face", StringComparison.OrdinalIgnoreCase)) {
                if (!TryValidateRasterFontFace(declarations, fontImports)) return true;
            } else {
                if (selector.StartsWith("@", StringComparison.Ordinal) || ContainsRasterDeclaration(body, RasterGeometryCssProperties)) return true;
                foreach (string item in SplitSvgCssTopLevel(selector, ',')) {
                    if (++ruleCount > MaximumSvgCssRules ||
                        !TryCalculateSvgSpecificity(item.Trim(), out _, out _)) return true;
                }
            }
            cursor = close + 1;
        }
        return false;
    }

    /// <summary>Validates every source through the canonical importer and its shared face, decoded-byte, retained-byte and work budgets.</summary>
    private static bool TryValidateRasterFontFace(IReadOnlyList<SvgCssDeclaration> declarations, SvgEmbeddedFontImports imports) {
        var body = new StringBuilder();
        foreach (SvgCssDeclaration declaration in declarations) {
            if (declaration.Important) return false;
            string value = declaration.Value;
            if (declaration.Name == "font-weight" && !OfficeFontFaceCssParser.TryWeight(value, 400, out _)) {
                string[] weights = value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (weights.Length != 2 ||
                    !int.TryParse(weights[0], NumberStyles.None, CultureInfo.InvariantCulture, out int minimum) ||
                    !int.TryParse(weights[1], NumberStyles.None, CultureInfo.InvariantCulture, out int maximum) ||
                    minimum < 1 || maximum > 1000 || minimum > maximum) return false;
                // A variable-face range is valid CSS but outside the editable single-weight importer.
                // Select a bounded weight only for validation; raster fallback receives the original SVG.
                value = Math.Max(minimum, Math.Min(maximum, 400)).ToString(CultureInfo.InvariantCulture);
            }
            body.Append(declaration.Name).Append(':').Append(value).Append(';');
        }
        int validatedDeclarations = 0;
        return TryImportSvgFontFace(body.ToString(), imports, ref validatedDeclarations);
    }

    private static bool IsBalancedRasterCssBlock(string value) {
        var delimiters = new Stack<char>();
        char quote = '\0';
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (current == '\\') {
                int escaped = index;
                if (!TryReadCssCharacter(value, ref escaped, out _)) return false;
                index = escaped - 1;
                continue;
            }
            if (quote != '\0') {
                if (current is '\r' or '\n' or '\f') return false;
                if (current == quote) quote = '\0';
            } else if (current is '\'' or '"') {
                quote = current;
            } else if (current is '{' or '}') {
                return false;
            } else if (current is '(' or '[') {
                if (delimiters.Count >= MaximumSvgNestingDepth) return false;
                delimiters.Push(current == '(' ? ')' : ']');
            } else if (current is ')' or ']') {
                if (delimiters.Count == 0 || delimiters.Pop() != current) return false;
            }
        }
        return quote == '\0' && delimiters.Count == 0;
    }
}
