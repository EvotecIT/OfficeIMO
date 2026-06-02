using System.Diagnostics;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.PowerPoint;

public sealed partial class OfficeMarkupPowerPointExporter {
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

    private static void AddPanel(PowerPointSlide slide, LayoutCursor box, OfficeMarkupResolvedStyle? style, string name) {
        var fillColor = ToPowerPointColor(style?.FillColor);
        var borderColor = ToPowerPointColor(style?.BorderColor);
        if (string.IsNullOrWhiteSpace(fillColor) && string.IsNullOrWhiteSpace(borderColor)) {
            return;
        }

        var panel = slide.AddShapeInches(A.ShapeTypeValues.Rectangle, box.Left, box.Top, box.Width, box.Height, name);
        if (!string.IsNullOrWhiteSpace(fillColor)) {
            panel.FillColor = fillColor;
        }

        if (!string.IsNullOrWhiteSpace(borderColor)) {
            panel.OutlineColor = borderColor;
            panel.OutlineWidthPoints = 0.75;
        }
    }
}
