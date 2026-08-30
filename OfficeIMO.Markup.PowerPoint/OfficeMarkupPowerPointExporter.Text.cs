using System.Diagnostics;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Markup.PowerPoint;

internal sealed partial class OfficeMarkupPowerPointExporter {
    private static bool IsMermaid(string language) =>
        string.Equals(language, "mermaid", StringComparison.OrdinalIgnoreCase);

    private static void ApplyTextStyle(PowerPointTextBox textBox, OfficeMarkupResolvedStyle? style) {
        textBox.SetTextMarginsInches(0.08, 0.04, 0.08, 0.04);

        if (style == null) {
            return;
        }

        if (!string.IsNullOrWhiteSpace(style.FontName)) {
            textBox.FontName = style.FontName;
        }

        if (style.FontSize != null) {
            textBox.FontSize = style.FontSize;
        }

        if (style.Bold != null) {
            textBox.Bold = style.Bold.Value;
        }

        if (style.Italic != null) {
            textBox.Italic = style.Italic.Value;
        }

        var textColor = ToPowerPointColor(style.TextColor);
        if (!string.IsNullOrWhiteSpace(textColor)) {
            textBox.Color = textColor;
        }

        foreach (PowerPointParagraph paragraph in textBox.Paragraphs) {
            foreach (PowerPointTextRun run in paragraph.Runs) {
                if (style.UnderlineStyle != null) run.UnderlineStyle = ToPowerPointUnderline(style.UnderlineStyle.Value);
                if (style.StrikethroughStyle != null) run.StrikeStyle = ToPowerPointStrike(style.StrikethroughStyle.Value);
                if (style.Baseline != null) ApplyBaseline(run, style.Baseline.Value);
                if (style.TextCase is { } textCase && textCase != OfficeTextCase.None) run.TransformTextCase(textCase, CultureInfo.InvariantCulture);
                if (style.SmallCaps != null) run.Capitalization = style.SmallCaps.Value ? PowerPointCapitalization.SmallCaps : PowerPointCapitalization.None;
                if (!string.IsNullOrWhiteSpace(style.HighlightColor)) run.HighlightColor = ToPowerPointColor(style.HighlightColor);
            }
        }

        var fillColor = ToPowerPointColor(style.FillColor);
        if (!string.IsNullOrWhiteSpace(fillColor)) {
            textBox.FillColor = fillColor;
        }

        var borderColor = ToPowerPointColor(style.BorderColor);
        if (!string.IsNullOrWhiteSpace(borderColor)) {
            textBox.OutlineColor = borderColor;
            textBox.OutlineWidthPoints = 0.75;
        }

        textBox.SetTextAutoFit(
            PowerPointTextAutoFit.Normal,
            new PowerPointTextAutoFitOptions(fontScalePercent: 82, lineSpaceReductionPercent: 18));
    }

    private static PowerPointUnderlineStyle ToPowerPointUnderline(OfficeTextDecorationStyle style) => style switch {
        OfficeTextDecorationStyle.None => PowerPointUnderlineStyle.None,
        OfficeTextDecorationStyle.Single => PowerPointUnderlineStyle.Single,
        OfficeTextDecorationStyle.Double => PowerPointUnderlineStyle.Double,
        OfficeTextDecorationStyle.Dotted => PowerPointUnderlineStyle.Dotted,
        OfficeTextDecorationStyle.Dashed => PowerPointUnderlineStyle.Dash,
        OfficeTextDecorationStyle.Wavy => PowerPointUnderlineStyle.Wavy,
        _ => throw new ArgumentOutOfRangeException(nameof(style))
    };

    private static PowerPointStrikeStyle ToPowerPointStrike(OfficeTextDecorationStyle style) => style switch {
        OfficeTextDecorationStyle.None => PowerPointStrikeStyle.None,
        OfficeTextDecorationStyle.Double => PowerPointStrikeStyle.Double,
        _ => PowerPointStrikeStyle.Single
    };

    private static void ApplyBaseline(PowerPointTextRun run, OfficeTextBaseline baseline) {
        switch (baseline) {
            case OfficeTextBaseline.Normal:
                run.SetBaseline();
                break;
            case OfficeTextBaseline.Superscript:
                run.SetSuperscript();
                break;
            case OfficeTextBaseline.Subscript:
                run.SetSubscript();
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(baseline));
        }
    }

    private static void AddPanel(PowerPointSlide slide, LayoutCursor box, OfficeMarkupResolvedStyle? style, string name) {
        var fillColor = ToPowerPointColor(style?.FillColor);
        var borderColor = ToPowerPointColor(style?.BorderColor);
        if (string.IsNullOrWhiteSpace(fillColor) && string.IsNullOrWhiteSpace(borderColor)) {
            return;
        }

        var panel = slide.AddShapeInches(OfficePresetShapeType.Rectangle, box.Left, box.Top, box.Width, box.Height, name);
        if (!string.IsNullOrWhiteSpace(fillColor)) {
            panel.FillColor = fillColor;
        }

        if (!string.IsNullOrWhiteSpace(borderColor)) {
            panel.OutlineColor = borderColor;
            panel.OutlineWidthPoints = 0.75;
        }
    }
}
