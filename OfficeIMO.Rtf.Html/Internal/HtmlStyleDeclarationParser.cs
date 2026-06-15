using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class HtmlStyleDeclarationParser {
    internal static HtmlStyleDeclaration Parse(string? style) {
        if (string.IsNullOrWhiteSpace(style)) {
            return HtmlStyleDeclaration.Empty;
        }

        var declaration = new HtmlStyleDeclaration();
        foreach (string rawDeclaration in SplitDeclarations(style!)) {
            int separator = rawDeclaration.IndexOf(':');
            if (separator <= 0) {
                continue;
            }

            string property = rawDeclaration.Substring(0, separator).Trim().ToLowerInvariant();
            string rawValue = NormalizeValue(rawDeclaration.Substring(separator + 1));
            string value = rawValue.ToLowerInvariant();
            if (value.Length == 0) {
                continue;
            }

            Apply(declaration, property, value, rawValue);
        }

        return declaration;
    }

    private static void Apply(HtmlStyleDeclaration declaration, string property, string value, string rawValue) {
        switch (property) {
            case "font-weight":
                declaration.Bold = ParseFontWeight(value);
                break;
            case "font-style":
                declaration.Italic = ParseFontStyle(value);
                break;
            case "font-family":
                declaration.FontFamily = ParseFontFamily(rawValue);
                break;
            case "font-size":
                declaration.FontSizePoints = ParseFontSize(value);
                break;
            case "text-decoration":
            case "text-decoration-line":
                ApplyTextDecoration(declaration, value);
                break;
            case "text-decoration-style":
                ApplyTextDecorationStyle(declaration, value);
                break;
            case "text-decoration-color":
                declaration.UnderlineColor = ParseColor(value);
                break;
            case "--officeimo-rtf-underline-style":
                declaration.UnderlineStyle = ParseRtfUnderlineStyle(value);
                break;
            case "--officeimo-rtf-strike-style":
                declaration.DoubleStrike = ParseRtfStrikeStyle(value);
                break;
            case "vertical-align":
                declaration.VerticalPosition = ParseVerticalAlign(value);
                declaration.TableCellVerticalAlignment = ParseTableCellVerticalAlign(value);
                break;
            case "text-align":
                declaration.TextAlignment = ParseTextAlign(value);
                break;
            case "width":
                if (TryParseTableWidth(value, out int width, out RtfTableWidthUnit widthUnit)) {
                    declaration.TableWidth = width;
                    declaration.TableWidthUnit = widthUnit;
                }

                break;
            case "height":
                if (TryParseTwips(value, out int heightTwips)) {
                    declaration.TableHeightTwips = heightTwips;
                }

                break;
            case "white-space":
                declaration.NoWrap = ParseWhiteSpace(value);
                break;
            case "padding":
                ApplyPadding(declaration, value);
                break;
            case "padding-top":
                declaration.PaddingTopTwips = ParseTwips(value);
                break;
            case "padding-left":
                declaration.PaddingLeftTwips = ParseTwips(value);
                break;
            case "padding-bottom":
                declaration.PaddingBottomTwips = ParseTwips(value);
                break;
            case "padding-right":
                declaration.PaddingRightTwips = ParseTwips(value);
                break;
            case "border":
                HtmlBorderDeclaration? border = ParseBorder(value);
                if (border != null) {
                    declaration.TopBorder = CloneBorder(border);
                    declaration.LeftBorder = CloneBorder(border);
                    declaration.BottomBorder = CloneBorder(border);
                    declaration.RightBorder = CloneBorder(border);
                }

                break;
            case "border-top":
                declaration.TopBorder = ParseBorder(value);
                break;
            case "border-left":
                declaration.LeftBorder = ParseBorder(value);
                break;
            case "border-bottom":
                declaration.BottomBorder = ParseBorder(value);
                break;
            case "border-right":
                declaration.RightBorder = ParseBorder(value);
                break;
            case "margin-left":
                declaration.LeftIndentTwips = ParseTwips(value);
                break;
            case "margin-right":
                declaration.RightIndentTwips = ParseTwips(value);
                break;
            case "margin-top":
                declaration.SpaceBeforeTwips = ParseTwips(value);
                break;
            case "margin-bottom":
                declaration.SpaceAfterTwips = ParseTwips(value);
                break;
            case "text-indent":
                declaration.FirstLineIndentTwips = ParseTwips(value);
                break;
            case "line-height":
                ApplyLineHeight(declaration, value);
                break;
            case "page-break-before":
            case "break-before":
                declaration.PageBreakBefore = IsPageBreakValue(value);
                break;
            case "page-break-after":
            case "break-after":
                declaration.PageBreakAfter = IsPageBreakValue(value);
                break;
            case "color":
                declaration.ForegroundColor = ParseColor(value);
                break;
            case "background":
            case "background-color":
                declaration.BackgroundColor = ParseColor(value);
                break;
        }
    }

    private static void ApplyLineHeight(HtmlStyleDeclaration declaration, string value) {
        if (value == "normal") {
            return;
        }

        if (value.EndsWith("%", StringComparison.Ordinal) &&
            double.TryParse(value.Substring(0, value.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent) &&
            percent > 0) {
            declaration.LineSpacingTwips = (int)Math.Round(percent * 240d / 100d, MidpointRounding.AwayFromZero);
            declaration.LineSpacingMultiple = true;
            return;
        }

        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double multiple) &&
            multiple > 0) {
            declaration.LineSpacingTwips = (int)Math.Round(multiple * 240d, MidpointRounding.AwayFromZero);
            declaration.LineSpacingMultiple = true;
            return;
        }

        int? twips = ParseTwips(value);
        if (twips.HasValue) {
            declaration.LineSpacingTwips = twips.Value;
            declaration.LineSpacingMultiple = false;
        }
    }

    private static void ApplyPadding(HtmlStyleDeclaration declaration, string value) {
        string[] parts = value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0) {
            return;
        }

        int? top = ParseTwips(parts[0]);
        int? right = ParseTwips(parts.Length > 1 ? parts[1] : parts[0]);
        int? bottom = ParseTwips(parts.Length > 2 ? parts[2] : parts[0]);
        int? left = ParseTwips(parts.Length > 3 ? parts[3] : (parts.Length > 1 ? parts[1] : parts[0]));
        declaration.PaddingTopTwips = top;
        declaration.PaddingRightTwips = right;
        declaration.PaddingBottomTwips = bottom;
        declaration.PaddingLeftTwips = left;
    }

    private static bool? ParseWhiteSpace(string value) {
        switch (value) {
            case "nowrap":
            case "pre":
            case "pre-line":
            case "pre-wrap":
                return true;
            case "normal":
                return false;
            default:
                return null;
        }
    }

    internal static bool TryParseTableWidth(string value, out int width, out RtfTableWidthUnit unit) {
        string normalized = NormalizeValue(value).Trim().ToLowerInvariant();
        width = 0;
        unit = RtfTableWidthUnit.Twips;

        if (normalized == "auto") {
            unit = RtfTableWidthUnit.Auto;
            return true;
        }

        if (normalized.EndsWith("%", StringComparison.Ordinal) &&
            double.TryParse(normalized.Substring(0, normalized.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent) &&
            percent > 0) {
            width = (int)Math.Round(percent * 50d, MidpointRounding.AwayFromZero);
            unit = RtfTableWidthUnit.Percent;
            return true;
        }

        int? twips = ParseTwips(normalized);
        if (twips.HasValue && twips.Value > 0) {
            width = twips.Value;
            unit = RtfTableWidthUnit.Twips;
            return true;
        }

        return false;
    }

    internal static bool TryParseTwips(string value, out int twips) {
        int? parsed = ParseTwips(NormalizeValue(value).Trim().ToLowerInvariant());
        twips = parsed.GetValueOrDefault();
        return parsed.HasValue;
    }

    internal static bool TryParseColor(string value, out RtfColor? color) {
        color = ParseColor(NormalizeValue(value).Trim().ToLowerInvariant());
        return color != null;
    }

    private static IEnumerable<string> SplitDeclarations(string style) {
        int start = 0;
        char quote = '\0';
        int parentheses = 0;
        for (int index = 0; index < style.Length; index++) {
            char current = style[index];
            if (quote != '\0') {
                if (current == quote) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
            } else if (current == '(') {
                parentheses++;
            } else if (current == ')' && parentheses > 0) {
                parentheses--;
            } else if (current == ';' && parentheses == 0) {
                yield return style.Substring(start, index - start);
                start = index + 1;
            }
        }

        if (start < style.Length) {
            yield return style.Substring(start);
        }
    }

    private static string NormalizeValue(string value) {
        string normalized = value.Trim();
        int important = normalized.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
        return important < 0 ? normalized : normalized.Substring(0, important).TrimEnd();
    }

    private static bool? ParseFontWeight(string value) {
        if (value == "bold" || value == "bolder") {
            return true;
        }

        if (value == "normal" || value == "lighter") {
            return false;
        }

        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int weight)) {
            return weight >= 600;
        }

        return null;
    }

    private static bool? ParseFontStyle(string value) {
        if (value == "italic" || value == "oblique") {
            return true;
        }

        if (value == "normal") {
            return false;
        }

        return null;
    }

    private static string? ParseFontFamily(string value) {
        foreach (string candidate in SplitFontFamily(value)) {
            string family = Unquote(candidate.Trim());
            if (family.Length == 0 || IsGenericFontFamily(family)) {
                continue;
            }

            return family;
        }

        return null;
    }

    private static IEnumerable<string> SplitFontFamily(string value) {
        int start = 0;
        char quote = '\0';
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (quote != '\0') {
                if (current == quote) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
            } else if (current == ',') {
                yield return value.Substring(start, index - start);
                start = index + 1;
            }
        }

        if (start < value.Length) {
            yield return value.Substring(start);
        }
    }

    private static string Unquote(string value) {
        if (value.Length >= 2 &&
            ((value[0] == '"' && value[value.Length - 1] == '"') ||
             (value[0] == '\'' && value[value.Length - 1] == '\''))) {
            return value.Substring(1, value.Length - 2).Trim();
        }

        return value;
    }

    private static bool IsGenericFontFamily(string value) {
        switch (value.Trim().ToLowerInvariant()) {
            case "serif":
            case "sans-serif":
            case "monospace":
            case "cursive":
            case "fantasy":
            case "system-ui":
            case "ui-serif":
            case "ui-sans-serif":
            case "ui-monospace":
                return true;
            default:
                return false;
        }
    }

    private static double? ParseFontSize(string value) {
        switch (value) {
            case "xx-small":
                return 6.75d;
            case "x-small":
                return 7.5d;
            case "small":
                return 10d;
            case "medium":
                return 12d;
            case "large":
                return 13.5d;
            case "x-large":
                return 18d;
            case "xx-large":
                return 24d;
            case "smaller":
                return 10d;
            case "larger":
                return 14d;
        }

        if (TryParseCssLength(value, "pt", 1d, out double points) ||
            TryParseCssLength(value, "px", 0.75d, out points) ||
            TryParseCssLength(value, "pc", 12d, out points) ||
            TryParseCssLength(value, "in", 72d, out points) ||
            TryParseCssLength(value, "cm", 72d / 2.54d, out points) ||
            TryParseCssLength(value, "mm", 72d / 25.4d, out points)) {
            return points > 0 ? points : null;
        }

        if (value.EndsWith("%", StringComparison.Ordinal) &&
            double.TryParse(value.Substring(0, value.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent) &&
            percent > 0) {
            return 12d * percent / 100d;
        }

        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double px) && px > 0
            ? px * 0.75d
            : null;
    }

    private static int? ParseTwips(string value) {
        if (TryParseCssLength(value, "pt", 1d, out double points) ||
            TryParseCssLength(value, "px", 0.75d, out points) ||
            TryParseCssLength(value, "pc", 12d, out points) ||
            TryParseCssLength(value, "in", 72d, out points) ||
            TryParseCssLength(value, "cm", 72d / 2.54d, out points) ||
            TryParseCssLength(value, "mm", 72d / 25.4d, out points)) {
            return (int)Math.Round(points * 20d, MidpointRounding.AwayFromZero);
        }

        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double px)
            ? (int)Math.Round(px * 15d, MidpointRounding.AwayFromZero)
            : null;
    }

    private static bool IsPageBreakValue(string value) {
        switch (value) {
            case "always":
            case "page":
            case "left":
            case "right":
            case "recto":
            case "verso":
                return true;
            default:
                return false;
        }
    }

    private static bool TryParseCssLength(string value, string unit, double multiplier, out double points) {
        points = 0;
        if (!value.EndsWith(unit, StringComparison.Ordinal)) {
            return false;
        }

        string number = value.Substring(0, value.Length - unit.Length).Trim();
        if (!double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) {
            return false;
        }

        points = parsed * multiplier;
        return true;
    }

    private static RtfVerticalPosition? ParseVerticalAlign(string value) {
        switch (value) {
            case "super":
            case "text-top":
                return RtfVerticalPosition.Superscript;
            case "sub":
            case "text-bottom":
                return RtfVerticalPosition.Subscript;
            case "baseline":
            case "middle":
                return RtfVerticalPosition.Baseline;
            default:
                return null;
        }
    }

    private static RtfTableCellVerticalAlignment? ParseTableCellVerticalAlign(string value) {
        switch (value) {
            case "top":
            case "text-top":
                return RtfTableCellVerticalAlignment.Top;
            case "middle":
                return RtfTableCellVerticalAlignment.Center;
            case "bottom":
            case "text-bottom":
                return RtfTableCellVerticalAlignment.Bottom;
            default:
                return null;
        }
    }

    private static RtfTextAlignment? ParseTextAlign(string value) {
        switch (value) {
            case "left":
            case "start":
                return RtfTextAlignment.Left;
            case "center":
                return RtfTextAlignment.Center;
            case "right":
            case "end":
                return RtfTextAlignment.Right;
            case "justify":
                return RtfTextAlignment.Justify;
            default:
                return null;
        }
    }

    private static RtfColor? ParseColor(string value) {
        if (value == "transparent" || value == "inherit" || value == "initial" || value == "currentcolor") {
            return null;
        }

        bool isRgbFunction = value.StartsWith("rgb(", StringComparison.Ordinal) || value.StartsWith("rgba(", StringComparison.Ordinal);
        int whitespace = value.IndexOfAny(new[] { ' ', '\t', '\r', '\n' });
        string token = whitespace > 0 && !isRgbFunction ? value.Substring(0, whitespace) : value;
        if (TryParseHexColor(token, out RtfColor? hexColor)) {
            return hexColor;
        }

        if (TryParseRgbColor(value, out RtfColor? rgbColor)) {
            return rgbColor;
        }

        return TryParseNamedColor(token, out RtfColor? namedColor) ? namedColor : null;
    }

    private static bool TryParseHexColor(string value, out RtfColor? color) {
        color = null;
        if (!value.StartsWith("#", StringComparison.Ordinal) || (value.Length != 4 && value.Length != 7)) {
            return false;
        }

        if (value.Length == 4) {
            if (!TryParseHexNibble(value[1], out byte r) ||
                !TryParseHexNibble(value[2], out byte g) ||
                !TryParseHexNibble(value[3], out byte b)) {
                return false;
            }

            color = new RtfColor((byte)(r * 17), (byte)(g * 17), (byte)(b * 17));
            return true;
        }

        if (!byte.TryParse(value.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte red) ||
            !byte.TryParse(value.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte green) ||
            !byte.TryParse(value.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte blue)) {
            return false;
        }

        color = new RtfColor(red, green, blue);
        return true;
    }

    private static bool TryParseRgbColor(string value, out RtfColor? color) {
        color = null;
        int open = value.IndexOf('(');
        int close = value.LastIndexOf(')');
        if (open < 0 || close <= open) {
            return false;
        }

        string function = value.Substring(0, open).Trim();
        if (function != "rgb" && function != "rgba") {
            return false;
        }

        string[] parts = value.Substring(open + 1, close - open - 1).Split(',');
        if (parts.Length < 3 ||
            !TryParseCssByte(parts[0], out byte red) ||
            !TryParseCssByte(parts[1], out byte green) ||
            !TryParseCssByte(parts[2], out byte blue)) {
            return false;
        }

        color = new RtfColor(red, green, blue);
        return true;
    }

    private static bool TryParseCssByte(string value, out byte component) {
        string normalized = value.Trim();
        if (normalized.EndsWith("%", StringComparison.Ordinal)) {
            if (double.TryParse(normalized.Substring(0, normalized.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent)) {
                component = ClampByte((int)Math.Round(percent * 2.55d, MidpointRounding.AwayFromZero));
                return true;
            }

            component = 0;
            return false;
        }

        if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
            component = ClampByte((int)Math.Round(number, MidpointRounding.AwayFromZero));
            return true;
        }

        component = 0;
        return false;
    }

    private static bool TryParseNamedColor(string value, out RtfColor? color) {
        switch (value) {
            case "black":
                color = new RtfColor(0, 0, 0);
                return true;
            case "white":
                color = new RtfColor(255, 255, 255);
                return true;
            case "red":
                color = new RtfColor(255, 0, 0);
                return true;
            case "green":
                color = new RtfColor(0, 128, 0);
                return true;
            case "blue":
                color = new RtfColor(0, 0, 255);
                return true;
            case "yellow":
                color = new RtfColor(255, 255, 0);
                return true;
            case "orange":
                color = new RtfColor(255, 165, 0);
                return true;
            case "purple":
                color = new RtfColor(128, 0, 128);
                return true;
            case "gray":
            case "grey":
                color = new RtfColor(128, 128, 128);
                return true;
            default:
                color = null;
                return false;
        }
    }

    private static bool TryParseHexNibble(char value, out byte result) {
        if (value >= '0' && value <= '9') {
            result = (byte)(value - '0');
            return true;
        }

        if (value >= 'a' && value <= 'f') {
            result = (byte)(value - 'a' + 10);
            return true;
        }

        if (value >= 'A' && value <= 'F') {
            result = (byte)(value - 'A' + 10);
            return true;
        }

        result = 0;
        return false;
    }

    private static byte ClampByte(int value) {
        return (byte)Math.Max(0, Math.Min(255, value));
    }

    private static bool ContainsWord(string value, string word) {
        return value.IndexOf(" " + word + " ", StringComparison.Ordinal) >= 0;
    }
}
