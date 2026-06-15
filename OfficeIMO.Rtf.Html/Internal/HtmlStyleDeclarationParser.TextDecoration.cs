namespace OfficeIMO.Rtf.Html;

internal static partial class HtmlStyleDeclarationParser {
    private static void ApplyTextDecoration(HtmlStyleDeclaration declaration, string value) {
        string normalized = " " + value.Replace('-', ' ') + " ";
        if (ContainsWord(normalized, "none")) {
            declaration.Underline = false;
            declaration.UnderlineStyle = RtfUnderlineStyle.None;
            declaration.UnderlineColor = null;
            declaration.Strike = false;
            return;
        }

        if (ContainsWord(normalized, "underline")) {
            declaration.Underline = true;
        }

        if (ContainsWord(normalized, "line through")) {
            declaration.Strike = true;
        }

        foreach (string part in SplitTextDecorationParts(value)) {
            RtfUnderlineStyle? style = ParseTextDecorationStyle(part);
            if (style.HasValue) {
                declaration.UnderlineStyle = style;
                continue;
            }

            RtfColor? color = ParseColor(part);
            if (color != null) {
                declaration.UnderlineColor = color;
            }
        }
    }

    private static RtfUnderlineStyle? ParseTextDecorationStyle(string value) {
        switch (value) {
            case "solid":
                return RtfUnderlineStyle.Single;
            case "double":
                return RtfUnderlineStyle.Double;
            case "dotted":
                return RtfUnderlineStyle.Dotted;
            case "dashed":
                return RtfUnderlineStyle.Dash;
            case "wavy":
                return RtfUnderlineStyle.Wave;
            default:
                return null;
        }
    }

    private static RtfUnderlineStyle? ParseRtfUnderlineStyle(string value) {
        switch (value) {
            case "none":
                return RtfUnderlineStyle.None;
            case "single":
                return RtfUnderlineStyle.Single;
            case "words":
                return RtfUnderlineStyle.Words;
            case "double":
                return RtfUnderlineStyle.Double;
            case "dotted":
                return RtfUnderlineStyle.Dotted;
            case "dash":
                return RtfUnderlineStyle.Dash;
            case "dash-dot":
                return RtfUnderlineStyle.DashDot;
            case "dash-dot-dot":
                return RtfUnderlineStyle.DashDotDot;
            case "thick":
                return RtfUnderlineStyle.Thick;
            case "thick-dotted":
                return RtfUnderlineStyle.ThickDotted;
            case "thick-dash":
                return RtfUnderlineStyle.ThickDash;
            case "thick-dash-dot":
                return RtfUnderlineStyle.ThickDashDot;
            case "thick-dash-dot-dot":
                return RtfUnderlineStyle.ThickDashDotDot;
            case "wave":
                return RtfUnderlineStyle.Wave;
            case "heavy-wave":
                return RtfUnderlineStyle.HeavyWave;
            case "double-wave":
                return RtfUnderlineStyle.DoubleWave;
            case "long-dash":
                return RtfUnderlineStyle.LongDash;
            case "thick-long-dash":
                return RtfUnderlineStyle.ThickLongDash;
            default:
                return null;
        }
    }

    private static string[] SplitTextDecorationParts(string value) =>
        value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
}
