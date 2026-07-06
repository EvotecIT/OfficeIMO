using System.Globalization;
using System.Net;

namespace OfficeIMO.Pdf;

internal readonly struct PdfFreeTextDefaultStyle {
    public PdfFreeTextDefaultStyle(double? fontSize, PdfColor? textColor, PdfAlign? textAlign) {
        FontSize = fontSize;
        TextColor = textColor;
        TextAlign = textAlign;
    }

    public double? FontSize { get; }

    public PdfColor? TextColor { get; }

    public PdfAlign? TextAlign { get; }
}

internal readonly struct PdfFreeTextRichTextRun {
    public PdfFreeTextRichTextRun(string text, bool bold, bool italic, bool underline, bool strike, PdfColor? color, double? fontSize, bool isLineBreak) {
        Text = text;
        Bold = bold;
        Italic = italic;
        Underline = underline;
        Strike = strike;
        Color = color;
        FontSize = fontSize;
        IsLineBreak = isLineBreak;
    }

    public string Text { get; }

    public bool Bold { get; }

    public bool Italic { get; }

    public bool Underline { get; }

    public bool Strike { get; }

    public PdfColor? Color { get; }

    public double? FontSize { get; }

    public bool IsLineBreak { get; }
}

internal readonly struct PdfFreeTextRichTextStyle {
    public PdfFreeTextRichTextStyle(bool bold, bool italic, bool underline, bool strike, PdfColor? color, double? fontSize) {
        Bold = bold;
        Italic = italic;
        Underline = underline;
        Strike = strike;
        Color = color;
        FontSize = fontSize;
    }

    public bool Bold { get; }

    public bool Italic { get; }

    public bool Underline { get; }

    public bool Strike { get; }

    public PdfColor? Color { get; }

    public double? FontSize { get; }
}

internal static class PdfFreeTextStyleParser {
    private static readonly char[] DeclarationSeparators = { ';' };
    private static readonly char[] RgbSeparators = { ',' };

    public static PdfFreeTextDefaultStyle ParseDefaultStyle(string? defaultStyle) {
        if (string.IsNullOrWhiteSpace(defaultStyle)) {
            return new PdfFreeTextDefaultStyle(null, null, null);
        }

        double? fontSize = null;
        PdfColor? textColor = null;
        PdfAlign? textAlign = null;
        string[] declarations = defaultStyle!.Split(DeclarationSeparators, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < declarations.Length; i++) {
            string declaration = declarations[i];
            int separator = declaration.IndexOf(':');
            if (separator <= 0 || separator >= declaration.Length - 1) {
                continue;
            }

            string property = declaration.Substring(0, separator).Trim();
            string value = declaration.Substring(separator + 1).Trim();
            if (property.Length == 0 || value.Length == 0) {
                continue;
            }

            if (string.Equals(property, "font-size", StringComparison.OrdinalIgnoreCase) &&
                TryReadCssFontSize(value, out double parsedFontSize)) {
                fontSize = parsedFontSize;
                continue;
            }

            if (string.Equals(property, "font", StringComparison.OrdinalIgnoreCase) &&
                TryReadCssFontSize(value, out parsedFontSize)) {
                fontSize = parsedFontSize;
                continue;
            }

            if (string.Equals(property, "color", StringComparison.OrdinalIgnoreCase) &&
                TryReadCssColor(value, out PdfColor parsedColor)) {
                textColor = parsedColor;
                continue;
            }

            if (string.Equals(property, "text-align", StringComparison.OrdinalIgnoreCase) &&
                TryReadCssTextAlign(value, out PdfAlign parsedAlign)) {
                textAlign = parsedAlign;
            }
        }

        return new PdfFreeTextDefaultStyle(fontSize, textColor, textAlign);
    }

    public static string? ExtractPlainText(string? richContents) {
        if (string.IsNullOrWhiteSpace(richContents)) {
            return null;
        }

        var builder = new System.Text.StringBuilder(richContents!.Length);
        for (int i = 0; i < richContents.Length; i++) {
            char current = richContents[i];
            if (current != '<') {
                builder.Append(current);
                continue;
            }

            int tagEnd = richContents.IndexOf('>', i + 1);
            if (tagEnd < 0) {
                builder.Append(current);
                continue;
            }

            AppendLineBreakForTag(builder, richContents, i + 1, tagEnd);
            i = tagEnd;
        }

        string decoded = WebUtility.HtmlDecode(builder.ToString());
        string normalized = NormalizeExtractedText(decoded);
        return normalized.Length == 0 ? null : normalized;
    }

    public static IReadOnlyList<PdfFreeTextRichTextRun>? ExtractRichTextRuns(string? richContents) {
        if (string.IsNullOrWhiteSpace(richContents)) {
            return null;
        }

        var runs = new List<PdfFreeTextRichTextRun>();
        var styles = new Stack<PdfFreeTextRichTextStyle>();
        styles.Push(new PdfFreeTextRichTextStyle(false, false, false, false, null, null));

        int textStart = 0;
        for (int i = 0; i < richContents!.Length; i++) {
            if (richContents[i] != '<') {
                continue;
            }

            int tagEnd = richContents.IndexOf('>', i + 1);
            if (tagEnd < 0) {
                continue;
            }

            AppendRichText(runs, richContents.Substring(textStart, i - textStart), styles.Peek());
            HandleRichTextTag(runs, styles, richContents, i + 1, tagEnd);
            i = tagEnd;
            textStart = tagEnd + 1;
        }

        if (textStart < richContents.Length) {
            AppendRichText(runs, richContents.Substring(textStart), styles.Peek());
        }

        TrimRichTextLineBreaks(runs);
        return HasRichText(runs) ? runs : null;
    }

    private static bool TryReadCssFontSize(string value, out double fontSize) {
        fontSize = 0D;
        for (int i = 0; i < value.Length; i++) {
            if (!char.IsDigit(value[i]) && value[i] != '.') {
                continue;
            }

            int start = i;
            i++;
            while (i < value.Length && (char.IsDigit(value[i]) || value[i] == '.')) {
                i++;
            }

            string number = value.Substring(start, i - start);
            if (double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) &&
                parsed > 0D &&
                !double.IsNaN(parsed) &&
                !double.IsInfinity(parsed)) {
                fontSize = parsed;
                return true;
            }
        }

        return false;
    }

    private static bool TryReadCssColor(string value, out PdfColor color) {
        value = value.Trim();
        if (TryReadHexColor(value, out color) ||
            TryReadRgbColor(value, out color) ||
            TryReadNamedColor(value, out color)) {
            return true;
        }

        color = PdfColor.Black;
        return false;
    }

    private static bool TryReadHexColor(string value, out PdfColor color) {
        color = PdfColor.Black;
        if (value.Length == 4 && value[0] == '#') {
            return TryReadHexDigit(value[1], out int r) &&
                TryReadHexDigit(value[2], out int g) &&
                TryReadHexDigit(value[3], out int b) &&
                TryCreateRgbColor(r * 17, g * 17, b * 17, out color);
        }

        if (value.Length == 7 && value[0] == '#') {
            return TryReadHexByte(value, 1, out int r) &&
                TryReadHexByte(value, 3, out int g) &&
                TryReadHexByte(value, 5, out int b) &&
                TryCreateRgbColor(r, g, b, out color);
        }

        return false;
    }

    private static bool TryReadRgbColor(string value, out PdfColor color) {
        color = PdfColor.Black;
        if (!value.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase) ||
            value.Length == 0 ||
            value[value.Length - 1] != ')') {
            return false;
        }

        string[] components = value.Substring(4, value.Length - 5).Split(RgbSeparators, StringSplitOptions.RemoveEmptyEntries);
        if (components.Length != 3 ||
            !TryReadRgbComponent(components[0], out double r) ||
            !TryReadRgbComponent(components[1], out double g) ||
            !TryReadRgbComponent(components[2], out double b)) {
            return false;
        }

        color = new PdfColor(r, g, b);
        return true;
    }

    private static bool TryReadNamedColor(string value, out PdfColor color) {
        switch (value.Trim().ToLowerInvariant()) {
            case "black":
                color = PdfColor.Black;
                return true;
            case "white":
                color = PdfColor.White;
                return true;
            case "red":
                color = new PdfColor(1D, 0D, 0D);
                return true;
            case "green":
                color = new PdfColor(0D, 0.5019607843137255D, 0D);
                return true;
            case "blue":
                color = new PdfColor(0D, 0D, 1D);
                return true;
            case "gray":
            case "grey":
                color = new PdfColor(0.5D, 0.5D, 0.5D);
                return true;
            case "yellow":
                color = new PdfColor(1D, 1D, 0D);
                return true;
            case "cyan":
                color = new PdfColor(0D, 1D, 1D);
                return true;
            case "magenta":
                color = new PdfColor(1D, 0D, 1D);
                return true;
            default:
                color = PdfColor.Black;
                return false;
        }
    }

    private static bool TryReadCssTextAlign(string value, out PdfAlign align) {
        switch (value.Trim().ToLowerInvariant()) {
            case "left":
            case "start":
                align = PdfAlign.Left;
                return true;
            case "center":
            case "middle":
                align = PdfAlign.Center;
                return true;
            case "right":
            case "end":
                align = PdfAlign.Right;
                return true;
            default:
                align = PdfAlign.Left;
                return false;
        }
    }

    private static void HandleRichTextTag(List<PdfFreeTextRichTextRun> runs, Stack<PdfFreeTextRichTextStyle> styles, string richContents, int tagStart, int tagEnd) {
        if (!TryReadTagName(richContents, tagStart, tagEnd, out string name, out bool closing, out bool selfClosing, out int attributesStart)) {
            return;
        }

        if (string.Equals(name, "br", StringComparison.OrdinalIgnoreCase)) {
            AppendRichLineBreak(runs);
            return;
        }

        if (string.Equals(name, "p", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "div", StringComparison.OrdinalIgnoreCase)) {
            AppendRichLineBreak(runs);
            return;
        }

        if (!IsRichStyleTag(name)) {
            return;
        }

        if (closing) {
            if (styles.Count > 1) {
                styles.Pop();
            }

            return;
        }

        PdfFreeTextRichTextStyle next = ApplyRichTextTagStyle(styles.Peek(), name);
        if (string.Equals(name, "span", StringComparison.OrdinalIgnoreCase)) {
            ApplySpanStyleAttributes(richContents, attributesStart, tagEnd, ref next);
        }

        styles.Push(next);
        if (selfClosing && styles.Count > 1) {
            styles.Pop();
        }
    }

    private static bool TryReadTagName(string value, int tagStart, int tagEnd, out string name, out bool closing, out bool selfClosing, out int attributesStart) {
        int index = tagStart;
        while (index < tagEnd && char.IsWhiteSpace(value[index])) {
            index++;
        }

        closing = index < tagEnd && value[index] == '/';
        if (closing) {
            index++;
        }

        while (index < tagEnd && char.IsWhiteSpace(value[index])) {
            index++;
        }

        int nameStart = index;
        while (index < tagEnd && (char.IsLetterOrDigit(value[index]) || value[index] == '-' || value[index] == ':')) {
            index++;
        }

        if (index <= nameStart) {
            name = string.Empty;
            selfClosing = false;
            attributesStart = tagEnd;
            return false;
        }

        name = value.Substring(nameStart, index - nameStart);
        attributesStart = index;
        int end = tagEnd - 1;
        while (end >= tagStart && char.IsWhiteSpace(value[end])) {
            end--;
        }

        selfClosing = end >= tagStart && value[end] == '/';
        return true;
    }

    private static bool IsRichStyleTag(string name) =>
        string.Equals(name, "b", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "strong", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "i", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "em", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "u", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "s", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "strike", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "span", StringComparison.OrdinalIgnoreCase);

    private static PdfFreeTextRichTextStyle ApplyRichTextTagStyle(PdfFreeTextRichTextStyle style, string name) {
        if (string.Equals(name, "b", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "strong", StringComparison.OrdinalIgnoreCase)) {
            return new PdfFreeTextRichTextStyle(true, style.Italic, style.Underline, style.Strike, style.Color, style.FontSize);
        }

        if (string.Equals(name, "i", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "em", StringComparison.OrdinalIgnoreCase)) {
            return new PdfFreeTextRichTextStyle(style.Bold, true, style.Underline, style.Strike, style.Color, style.FontSize);
        }

        if (string.Equals(name, "u", StringComparison.OrdinalIgnoreCase)) {
            return new PdfFreeTextRichTextStyle(style.Bold, style.Italic, true, style.Strike, style.Color, style.FontSize);
        }

        if (string.Equals(name, "s", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "strike", StringComparison.OrdinalIgnoreCase)) {
            return new PdfFreeTextRichTextStyle(style.Bold, style.Italic, style.Underline, true, style.Color, style.FontSize);
        }

        return style;
    }

    private static void ApplySpanStyleAttributes(string richContents, int attributesStart, int tagEnd, ref PdfFreeTextRichTextStyle style) {
        string? css = TryReadAttributeValue(richContents, attributesStart, tagEnd, "style");
        if (string.IsNullOrWhiteSpace(css)) {
            return;
        }

        string[] declarations = css!.Split(DeclarationSeparators, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < declarations.Length; i++) {
            string declaration = declarations[i];
            int separator = declaration.IndexOf(':');
            if (separator <= 0 || separator >= declaration.Length - 1) {
                continue;
            }

            string property = declaration.Substring(0, separator).Trim();
            string value = declaration.Substring(separator + 1).Trim();
            if (property.Length == 0 || value.Length == 0) {
                continue;
            }

            if (string.Equals(property, "font-weight", StringComparison.OrdinalIgnoreCase) && IsBoldCssValue(value)) {
                style = new PdfFreeTextRichTextStyle(true, style.Italic, style.Underline, style.Strike, style.Color, style.FontSize);
                continue;
            }

            if (string.Equals(property, "font-style", StringComparison.OrdinalIgnoreCase) && ContainsIgnoreCase(value, "italic")) {
                style = new PdfFreeTextRichTextStyle(style.Bold, true, style.Underline, style.Strike, style.Color, style.FontSize);
                continue;
            }

            if (string.Equals(property, "text-decoration", StringComparison.OrdinalIgnoreCase)) {
                bool underline = style.Underline || ContainsIgnoreCase(value, "underline");
                bool strike = style.Strike || ContainsIgnoreCase(value, "line-through");
                style = new PdfFreeTextRichTextStyle(style.Bold, style.Italic, underline, strike, style.Color, style.FontSize);
                continue;
            }

            if (string.Equals(property, "color", StringComparison.OrdinalIgnoreCase) && TryReadCssColor(value, out PdfColor color)) {
                style = new PdfFreeTextRichTextStyle(style.Bold, style.Italic, style.Underline, style.Strike, color, style.FontSize);
                continue;
            }

            if (string.Equals(property, "font-size", StringComparison.OrdinalIgnoreCase) && TryReadCssFontSize(value, out double fontSize)) {
                style = new PdfFreeTextRichTextStyle(style.Bold, style.Italic, style.Underline, style.Strike, style.Color, fontSize);
            }
        }
    }

    private static string? TryReadAttributeValue(string value, int attributesStart, int tagEnd, string attributeName) {
        int index = attributesStart;
        while (index < tagEnd) {
            while (index < tagEnd && char.IsWhiteSpace(value[index])) {
                index++;
            }

            int nameStart = index;
            while (index < tagEnd && (char.IsLetterOrDigit(value[index]) || value[index] == '-' || value[index] == ':')) {
                index++;
            }

            if (index <= nameStart) {
                index++;
                continue;
            }

            string name = value.Substring(nameStart, index - nameStart);
            while (index < tagEnd && char.IsWhiteSpace(value[index])) {
                index++;
            }

            if (index >= tagEnd || value[index] != '=') {
                continue;
            }

            index++;
            while (index < tagEnd && char.IsWhiteSpace(value[index])) {
                index++;
            }

            if (index >= tagEnd) {
                break;
            }

            char quote = value[index];
            bool quoted = quote == '"' || quote == '\'';
            if (quoted) {
                index++;
            }

            int valueStart = index;
            while (index < tagEnd && (quoted ? value[index] != quote : !char.IsWhiteSpace(value[index]) && value[index] != '/')) {
                index++;
            }

            string attributeValue = value.Substring(valueStart, index - valueStart);
            if (quoted && index < tagEnd) {
                index++;
            }

            if (string.Equals(name, attributeName, StringComparison.OrdinalIgnoreCase)) {
                return WebUtility.HtmlDecode(attributeValue);
            }
        }

        return null;
    }

    private static bool IsBoldCssValue(string value) {
        value = value.Trim();
        if (ContainsIgnoreCase(value, "bold")) {
            return true;
        }

        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int weight) && weight >= 600;
    }

    private static void AppendRichText(List<PdfFreeTextRichTextRun> runs, string encodedText, PdfFreeTextRichTextStyle style) {
        if (encodedText.Length == 0) {
            return;
        }

        string text = WebUtility.HtmlDecode(encodedText).Replace("\r\n", "\n").Replace('\r', '\n');
        var builder = new System.Text.StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++) {
            if (text[i] == '\n') {
                AppendRichTextRun(runs, builder.ToString(), style);
                builder.Clear();
                AppendRichLineBreak(runs);
                continue;
            }

            builder.Append(text[i]);
        }

        AppendRichTextRun(runs, builder.ToString(), style);
    }

    private static bool ContainsIgnoreCase(string source, string value) {
#if NET8_0_OR_GREATER
        return source.Contains(value, StringComparison.OrdinalIgnoreCase);
#else
        return source.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
#endif
    }

    private static void AppendRichTextRun(List<PdfFreeTextRichTextRun> runs, string text, PdfFreeTextRichTextStyle style) {
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        if (runs.Count > 0) {
            PdfFreeTextRichTextRun previous = runs[runs.Count - 1];
            if (!previous.IsLineBreak && HasSameStyle(previous, style)) {
                runs[runs.Count - 1] = new PdfFreeTextRichTextRun(
                    previous.Text + text,
                    previous.Bold,
                    previous.Italic,
                    previous.Underline,
                    previous.Strike,
                    previous.Color,
                    previous.FontSize,
                    isLineBreak: false);
                return;
            }
        }

        runs.Add(new PdfFreeTextRichTextRun(text, style.Bold, style.Italic, style.Underline, style.Strike, style.Color, style.FontSize, isLineBreak: false));
    }

    private static void AppendRichLineBreak(List<PdfFreeTextRichTextRun> runs) {
        if (runs.Count == 0 || runs[runs.Count - 1].IsLineBreak) {
            return;
        }

        runs.Add(new PdfFreeTextRichTextRun(string.Empty, false, false, false, false, null, null, isLineBreak: true));
    }

    private static void TrimRichTextLineBreaks(List<PdfFreeTextRichTextRun> runs) {
        while (runs.Count > 0 && runs[0].IsLineBreak) {
            runs.RemoveAt(0);
        }

        while (runs.Count > 0 && runs[runs.Count - 1].IsLineBreak) {
            runs.RemoveAt(runs.Count - 1);
        }
    }

    private static bool HasRichText(List<PdfFreeTextRichTextRun> runs) {
        for (int i = 0; i < runs.Count; i++) {
            if (!runs[i].IsLineBreak && runs[i].Text.Length > 0) {
                return true;
            }
        }

        return false;
    }

    private static bool HasSameStyle(PdfFreeTextRichTextRun run, PdfFreeTextRichTextStyle style) =>
        run.Bold == style.Bold &&
        run.Italic == style.Italic &&
        run.Underline == style.Underline &&
        run.Strike == style.Strike &&
        Nullable.Equals(run.Color, style.Color) &&
        Nullable.Equals(run.FontSize, style.FontSize);

    private static bool TryReadHexByte(string value, int offset, out int component) {
        component = 0;
        if (!TryReadHexDigit(value[offset], out int high) ||
            !TryReadHexDigit(value[offset + 1], out int low)) {
            return false;
        }

        component = high * 16 + low;
        return true;
    }

    private static bool TryReadHexDigit(char value, out int digit) {
        if (value >= '0' && value <= '9') {
            digit = value - '0';
            return true;
        }

        if (value >= 'a' && value <= 'f') {
            digit = value - 'a' + 10;
            return true;
        }

        if (value >= 'A' && value <= 'F') {
            digit = value - 'A' + 10;
            return true;
        }

        digit = 0;
        return false;
    }

    private static bool TryReadRgbComponent(string value, out double component) {
        value = value.Trim();
        bool isPercent = value.Length > 0 && value[value.Length - 1] == '%';
        string number = isPercent ? value.Substring(0, value.Length - 1) : value;
        if (!double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ||
            double.IsNaN(parsed) ||
            double.IsInfinity(parsed)) {
            component = 0D;
            return false;
        }

        component = Clamp01(isPercent ? parsed / 100D : parsed / 255D);
        return true;
    }

    private static bool TryCreateRgbColor(int r, int g, int b, out PdfColor color) {
        if (r < 0 || r > 255 || g < 0 || g > 255 || b < 0 || b > 255) {
            color = PdfColor.Black;
            return false;
        }

        color = new PdfColor(r / 255D, g / 255D, b / 255D);
        return true;
    }

    private static double Clamp01(double value) {
        if (value < 0D) {
            return 0D;
        }

        return value > 1D ? 1D : value;
    }

    private static void AppendLineBreakForTag(System.Text.StringBuilder builder, string richContents, int tagStart, int tagEnd) {
        int index = tagStart;
        while (index < tagEnd && char.IsWhiteSpace(richContents[index])) {
            index++;
        }

        bool closing = index < tagEnd && richContents[index] == '/';
        if (closing) {
            index++;
        }

        while (index < tagEnd && char.IsWhiteSpace(richContents[index])) {
            index++;
        }

        int nameStart = index;
        while (index < tagEnd && char.IsLetterOrDigit(richContents[index])) {
            index++;
        }

        if (index <= nameStart) {
            return;
        }

        string name = richContents.Substring(nameStart, index - nameStart);
        if (string.Equals(name, "br", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "p", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "div", StringComparison.OrdinalIgnoreCase)) {
            AppendSingleNewLine(builder, onlyWhenTextExists: !closing);
        }
    }

    private static void AppendSingleNewLine(System.Text.StringBuilder builder, bool onlyWhenTextExists) {
        if (onlyWhenTextExists && builder.Length == 0) {
            return;
        }

        if (builder.Length > 0 && builder[builder.Length - 1] == '\n') {
            return;
        }

        builder.Append('\n');
    }

    private static string NormalizeExtractedText(string value) {
        value = value.Replace("\r\n", "\n").Replace('\r', '\n').Trim();
        while (value.Contains("\n\n\n")) {
            value = value.Replace("\n\n\n", "\n\n");
        }

        return value;
    }
}
