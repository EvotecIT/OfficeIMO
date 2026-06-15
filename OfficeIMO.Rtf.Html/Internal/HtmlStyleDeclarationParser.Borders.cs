namespace OfficeIMO.Rtf.Html;

internal static partial class HtmlStyleDeclarationParser {
    private static HtmlBorderDeclaration? ParseBorder(string value) {
        string[] parts = value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0) {
            return null;
        }

        var border = new HtmlBorderDeclaration();
        bool hasStyle = false;
        bool hasWidth = false;
        bool hasColor = false;
        foreach (string part in parts) {
            if (!hasStyle && TryParseBorderStyle(part, out RtfTableCellBorderStyle style)) {
                border.Style = style;
                hasStyle = true;
                continue;
            }

            if (!hasWidth && TryParseBorderWidth(part, out int? width)) {
                border.Width = width;
                hasWidth = true;
                continue;
            }

            if (!hasColor && TryParseBorderColor(part, out RtfColor? color)) {
                border.Color = color;
                hasColor = true;
            }
        }

        return hasStyle || hasWidth || hasColor ? border : null;
    }

    private static HtmlBorderDeclaration CloneBorder(HtmlBorderDeclaration border) {
        return new HtmlBorderDeclaration {
            Style = border.Style,
            Width = border.Width,
            Color = border.Color
        };
    }

    private static bool TryParseBorderStyle(string value, out RtfTableCellBorderStyle style) {
        switch (value) {
            case "none":
            case "hidden":
                style = RtfTableCellBorderStyle.None;
                return true;
            case "double":
                style = RtfTableCellBorderStyle.Double;
                return true;
            case "dotted":
                style = RtfTableCellBorderStyle.Dotted;
                return true;
            case "dashed":
                style = RtfTableCellBorderStyle.Dashed;
                return true;
            case "solid":
                style = RtfTableCellBorderStyle.Single;
                return true;
            default:
                style = RtfTableCellBorderStyle.None;
                return false;
        }
    }

    private static bool TryParseBorderWidth(string value, out int? width) {
        switch (value) {
            case "thin":
                width = 10;
                return true;
            case "medium":
                width = 20;
                return true;
            case "thick":
                width = 30;
                return true;
        }

        width = ParseTwips(value);
        return width.HasValue;
    }

    private static bool TryParseBorderColor(string value, out RtfColor? color) {
        color = ParseColor(value);
        return color != null;
    }
}
