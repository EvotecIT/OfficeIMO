using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Internal;

/// <summary>
/// Turns an <see cref="OfficeTheme"/> into the CSS custom properties that HTML writers share:
/// <c>--office-bg</c>, <c>--office-text</c>, <c>--office-muted</c>, <c>--office-accent-1</c> to <c>-6</c>,
/// the status colours and their <c>-fill</c> tints, font stacks, the type scale in points and table tokens.
/// Every value comes from the theme; status fills are the status colour at 12 % over the background
/// (18 % for dark backgrounds), as the visual design derives them.
/// </summary>
internal static class OfficeThemeCssTokens {
    internal const string MonospaceStack = "ui-monospace, SFMono-Regular, Menlo, Consolas, \"Liberation Mono\", \"Courier New\", monospace";

    internal static string Build(OfficeTheme theme, string selector) {
        var css = new StringBuilder();
        css.Append(selector).Append(" {");
        if (IsDark(theme.Palette.Background)) css.Append(" color-scheme: dark;");
        foreach (KeyValuePair<string, string> token in GetTokens(theme)) {
            css.Append(' ').Append(token.Key).Append(": ").Append(token.Value).Append(';');
        }
        css.Append(" }\n");
        return css.ToString();
    }

    /// <summary>Replaces <c>var(--office-*)</c> references with the theme's literal values, for consumers without custom properties.</summary>
    internal static string ResolveReferences(string css, OfficeTheme theme) {
        string resolved = css;
        foreach (KeyValuePair<string, string> token in GetTokens(theme)) {
            resolved = resolved.Replace("var(" + token.Key + ")", token.Value);
        }
        return resolved;
    }

    internal static IReadOnlyList<KeyValuePair<string, string>> GetTokens(OfficeTheme theme) {
        if (theme == null) throw new ArgumentNullException(nameof(theme));
        OfficeThemePalette palette = theme.Palette;
        bool dark = IsDark(palette.Background);
        var tokens = new List<KeyValuePair<string, string>>();
        void Append(string name, string value) => tokens.Add(new KeyValuePair<string, string>(name, value));
        Append("--office-bg", Hex(palette.Background, theme));
        Append("--office-text", Hex(palette.Text, theme));
        Append("--office-muted", Hex(palette.TextMuted, theme));
        OfficeColor[] accents = { palette.Accent1, palette.Accent2, palette.Accent3, palette.Accent4, palette.Accent5, palette.Accent6 };
        // Longer names first so a reference to --office-accent-1-tint is never cut short by --office-accent-1.
        Append("--office-accent-1-tint", Mix(palette.Accent1, palette.Background, 0.1D, theme));
        for (int index = 0; index < accents.Length; index++) {
            Append("--office-accent-" + (index + 1).ToString(CultureInfo.InvariantCulture), Hex(accents[index], theme));
        }

        double tint = dark ? 0.18D : 0.12D;
        foreach ((string name, OfficeColor color) in new[] { ("success", palette.Success), ("warning", palette.Warning), ("danger", palette.Danger) }) {
            Append("--office-" + name + "-fill", Mix(color, palette.Background, tint, theme));
            Append("--office-" + name, Hex(color, theme));
        }

        Append("--office-font-heading", FontStack(theme.Fonts.Heading, "sans-serif"));
        Append("--office-font-body", FontStack(theme.Fonts.Body, "sans-serif"));
        Append("--office-font-mono", MonospaceStack);

        OfficeThemeTypeScale scale = theme.TypeScale;
        Append("--office-size-title", Points(scale.Title));
        Append("--office-size-h1", Points(scale.Heading1));
        Append("--office-size-h2", Points(scale.Heading2));
        Append("--office-size-h3", Points(scale.Heading3));
        Append("--office-size-h4", Points(scale.Heading4));
        Append("--office-size-h5", Points(scale.Heading5));
        Append("--office-size-h6", Points(scale.Heading6));
        Append("--office-size-body", Points(scale.Body));
        Append("--office-size-caption", Points(scale.Caption));
        Append("--office-paragraph-after", Points(theme.Spacing.ParagraphAfter));
        Append("--office-list-indent", Points(theme.Spacing.ListIndent));

        OfficeThemeTableStyle table = theme.TableStyle;
        Append("--office-table-header-fill", Hex(table.HeaderFill, theme));
        Append("--office-table-header-text", Hex(table.HeaderText, theme));
        Append("--office-table-band-fill", table.Banding ? Hex(table.BandFill, theme) : "transparent");
        Append("--office-table-border-width", Points(table.BorderWidth));
        Append("--office-table-border", Hex(table.BorderColor, theme));
        Append("--office-table-cell-padding-y", Points(table.CellPadding));
        Append("--office-table-cell-padding-x", Points(OfficeLength.Points(table.CellPadding.ToPoints() * 1.4D)));
        Append("--office-chart-gridline", Hex(theme.ChartStyle.GridlineColor, theme));
        return tokens;
    }

    /// <summary>Returns true when a table cell's text reads as a number: digits with optional sign, separators, decimals, currency or percent.</summary>
    internal static bool IsNumericText(string? value) {
        if (value == null) return false;
        string text = value.Trim();
        if (text.Length == 0 || text.Length > 40) return false;
        int start = 0;
        int end = text.Length;
        if (text[start] == '+' || text[start] == '-' || text[start] == '−' || text[start] == '(') start++;
        while (start < end && IsCurrency(text[start])) start++;
        while (end > start && (text[end - 1] == '%' || text[end - 1] == ')' || IsCurrency(text[end - 1]) || text[end - 1] == ' ')) end--;
        if (start >= end) return false;
        bool digit = false;
        bool decimalSeen = false;
        for (int index = start; index < end; index++) {
            char c = text[index];
            if (char.IsDigit(c)) { digit = true; continue; }
            if (c == '.' && !decimalSeen) { decimalSeen = true; continue; }
            if ((c == ',' || c == ' ' || c == ' ' || c == ' ' || c == '\'') && index > start && index + 1 < end && char.IsDigit(text[index + 1])) continue;
            return false;
        }
        return digit;
    }

    private static bool IsCurrency(char c) => c == '$' || c == '€' || c == '£' || c == '¥' || c == '₹' || c == '¢';

    private static string Hex(OfficeColor color, OfficeTheme theme) {
        OfficeColor resolved = color.IsThemeRole ? color.Resolve(theme) : color;
        return resolved.A == 0 ? "transparent" : "#" + resolved.ToRgbHex();
    }

    private static string Mix(OfficeColor color, OfficeColor background, double amount, OfficeTheme theme) {
        OfficeColor top = color.IsThemeRole ? color.Resolve(theme) : color;
        OfficeColor bottom = background.IsThemeRole ? background.Resolve(theme) : background;
        byte Channel(byte front, byte back) => (byte)Math.Round(front * amount + back * (1D - amount), MidpointRounding.AwayFromZero);
        return "#" + OfficeColor.FromRgb(Channel(top.R, bottom.R), Channel(top.G, bottom.G), Channel(top.B, bottom.B)).ToRgbHex();
    }

    private static bool IsDark(OfficeColor background) {
        if (background.IsThemeRole) return false;
        double luminance = (0.2126D * background.R + 0.7152D * background.G + 0.0722D * background.B) / 255D;
        return luminance < 0.4D;
    }

    private static string Points(OfficeLength length) =>
        length.ToPoints().ToString("0.##", CultureInfo.InvariantCulture) + "pt";

    /// <summary>Primary family, then its fallbacks, always ending in a generic family.</summary>
    internal static string FontStack(OfficeThemeFont font, string generic) {
        var names = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        void Add(string? name) {
            if (string.IsNullOrWhiteSpace(name)) return;
            string trimmed = name!.Trim().Trim('"', '\'');
            if (seen.Add(trimmed)) names.Add(trimmed);
        }

        Add(font.Family);
        foreach (string fallback in font.Fallbacks) Add(fallback);
        bool hasGeneric = false;
        foreach (string name in names) hasGeneric |= IsGeneric(name);
        if (!hasGeneric) names.Add(generic);
        var stack = new StringBuilder();
        foreach (string name in names) {
            if (stack.Length > 0) stack.Append(", ");
            stack.Append(IsGeneric(name) ? name : "\"" + name.Replace("\"", string.Empty) + "\"");
        }

        return stack.ToString();
    }

    private static bool IsGeneric(string name) => name is "serif" or "sans-serif" or "monospace" or "cursive" or "fantasy"
        or "system-ui" or "ui-serif" or "ui-sans-serif" or "ui-monospace" or "math";
}
