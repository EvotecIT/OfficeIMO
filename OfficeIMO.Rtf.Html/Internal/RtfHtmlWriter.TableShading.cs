using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendParagraphShadingStyle(StringBuilder builder, RtfParagraph paragraph, RtfDocument document) {
        AppendShadingForegroundStyle(builder, paragraph.ShadingForegroundColorIndex, document);
        AppendShadingIntegerStyle(builder, "--officeimo-rtf-shading-percent", paragraph.ShadingPatternPercent);
        AppendShadingPatternStyle(builder, paragraph.ShadingPattern);
    }

    private static void AppendCharacterShadingStyle(StringBuilder builder, RtfRun run, RtfDocument document) {
        AppendShadingForegroundStyle(builder, run.CharacterShadingForegroundColorIndex, document);
        AppendShadingIntegerStyle(builder, "--officeimo-rtf-shading-percent", run.CharacterShadingPatternPercent);
        AppendShadingPatternStyle(builder, run.CharacterShadingPattern);
    }

    private static void AppendTableRowShadingStyle(StringBuilder builder, RtfTableRow row, RtfDocument document) {
        AppendShadingForegroundStyle(builder, row.ShadingForegroundColorIndex, document);
        AppendShadingIntegerStyle(builder, "--officeimo-rtf-shading-pattern-value", row.ShadingPatternValue);
        AppendShadingIntegerStyle(builder, "--officeimo-rtf-shading-percent", row.ShadingPatternPercent);
        AppendShadingPatternStyle(builder, row.ShadingPattern);
    }

    private static void AppendTableCellShadingStyle(StringBuilder builder, RtfTableCell cell, RtfDocument document) {
        AppendShadingForegroundStyle(builder, cell.ShadingForegroundColorIndex, document);
        AppendShadingIntegerStyle(builder, "--officeimo-rtf-shading-percent", cell.ShadingPatternPercent);
        AppendShadingPatternStyle(builder, cell.ShadingPattern);
    }

    private static void AppendShadingForegroundStyle(StringBuilder builder, int? colorIndex, RtfDocument document) {
        if (!TryGetColor(document, colorIndex, out RtfColor? color)) {
            return;
        }

        builder.Append("--officeimo-rtf-shading-foreground:");
        builder.Append(FormatColor(color!));
        builder.Append(';');
    }

    private static void AppendShadingIntegerStyle(StringBuilder builder, string propertyName, int? value) {
        if (!value.HasValue) {
            return;
        }

        builder.Append(propertyName);
        builder.Append(':');
        builder.Append(value.Value.ToString(CultureInfo.InvariantCulture));
        builder.Append(';');
    }

    private static void AppendShadingPatternStyle(StringBuilder builder, RtfShadingPattern pattern) {
        if (pattern == RtfShadingPattern.None) {
            return;
        }

        builder.Append("--officeimo-rtf-shading-pattern:");
        builder.Append(FormatShadingPattern(pattern));
        builder.Append(';');
    }

    private static string FormatShadingPattern(RtfShadingPattern pattern) {
        switch (pattern) {
            case RtfShadingPattern.Horizontal:
                return "horizontal";
            case RtfShadingPattern.Vertical:
                return "vertical";
            case RtfShadingPattern.ForwardDiagonal:
                return "forward-diagonal";
            case RtfShadingPattern.BackwardDiagonal:
                return "backward-diagonal";
            case RtfShadingPattern.Cross:
                return "cross";
            case RtfShadingPattern.DiagonalCross:
                return "diagonal-cross";
            case RtfShadingPattern.DarkHorizontal:
                return "dark-horizontal";
            case RtfShadingPattern.DarkVertical:
                return "dark-vertical";
            case RtfShadingPattern.DarkForwardDiagonal:
                return "dark-forward-diagonal";
            case RtfShadingPattern.DarkBackwardDiagonal:
                return "dark-backward-diagonal";
            case RtfShadingPattern.DarkCross:
                return "dark-cross";
            case RtfShadingPattern.DarkDiagonalCross:
                return "dark-diagonal-cross";
            default:
                return "none";
        }
    }
}
