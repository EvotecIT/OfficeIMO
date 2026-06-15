using System.Text;

namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlWriter {
    private static bool HasRichUnderline(RtfRun run) =>
        run.Underline &&
        (run.UnderlineStyle != RtfUnderlineStyle.Single || run.UnderlineColorIndex.HasValue);

    private static bool HasRichStrike(RtfRun run) => run.DoubleStrike;

    private static void AppendTextDecorationStyle(StringBuilder builder, RtfRun run, RtfDocument document) {
        bool richUnderline = HasRichUnderline(run);
        bool richStrike = HasRichStrike(run);
        if (!richUnderline && !richStrike) {
            return;
        }

        builder.Append("text-decoration-line:");
        if (richUnderline) {
            builder.Append("underline");
        }

        if (richStrike) {
            if (richUnderline) {
                builder.Append(' ');
            }

            builder.Append("line-through");
        }

        builder.Append(';');
        builder.Append("text-decoration-style:");
        builder.Append(richUnderline ? FormatCssUnderlineStyle(run.UnderlineStyle) : "double");
        builder.Append(';');

        if (richUnderline) {
            builder.Append("--officeimo-rtf-underline-style:");
            builder.Append(FormatRtfUnderlineStyle(run.UnderlineStyle));
            builder.Append(';');
        }

        if (richStrike) {
            builder.Append("--officeimo-rtf-strike-style:double;");
        }

        if (TryGetColor(document, run.UnderlineColorIndex, out RtfColor? color)) {
            builder.Append("text-decoration-color:");
            builder.Append(FormatColor(color!));
            builder.Append(';');
        }
    }

    private static string FormatCssUnderlineStyle(RtfUnderlineStyle style) {
        switch (style) {
            case RtfUnderlineStyle.Double:
                return "double";
            case RtfUnderlineStyle.Dotted:
            case RtfUnderlineStyle.ThickDotted:
                return "dotted";
            case RtfUnderlineStyle.Dash:
            case RtfUnderlineStyle.DashDot:
            case RtfUnderlineStyle.DashDotDot:
            case RtfUnderlineStyle.ThickDash:
            case RtfUnderlineStyle.ThickDashDot:
            case RtfUnderlineStyle.ThickDashDotDot:
            case RtfUnderlineStyle.LongDash:
            case RtfUnderlineStyle.ThickLongDash:
                return "dashed";
            case RtfUnderlineStyle.Wave:
            case RtfUnderlineStyle.HeavyWave:
            case RtfUnderlineStyle.DoubleWave:
                return "wavy";
            default:
                return "solid";
        }
    }

    private static string FormatRtfUnderlineStyle(RtfUnderlineStyle style) {
        switch (style) {
            case RtfUnderlineStyle.None:
                return "none";
            case RtfUnderlineStyle.Words:
                return "words";
            case RtfUnderlineStyle.Double:
                return "double";
            case RtfUnderlineStyle.Dotted:
                return "dotted";
            case RtfUnderlineStyle.Dash:
                return "dash";
            case RtfUnderlineStyle.DashDot:
                return "dash-dot";
            case RtfUnderlineStyle.DashDotDot:
                return "dash-dot-dot";
            case RtfUnderlineStyle.Thick:
                return "thick";
            case RtfUnderlineStyle.ThickDotted:
                return "thick-dotted";
            case RtfUnderlineStyle.ThickDash:
                return "thick-dash";
            case RtfUnderlineStyle.ThickDashDot:
                return "thick-dash-dot";
            case RtfUnderlineStyle.ThickDashDotDot:
                return "thick-dash-dot-dot";
            case RtfUnderlineStyle.Wave:
                return "wave";
            case RtfUnderlineStyle.HeavyWave:
                return "heavy-wave";
            case RtfUnderlineStyle.DoubleWave:
                return "double-wave";
            case RtfUnderlineStyle.LongDash:
                return "long-dash";
            case RtfUnderlineStyle.ThickLongDash:
                return "thick-long-dash";
            default:
                return "single";
        }
    }
}
