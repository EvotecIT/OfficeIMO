using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        internal static OfficeDrawing CreateTextBoxDrawing(PowerPointTextBox textBox,
            double width, double height, List<OfficeImageExportDiagnostic> diagnostics,
            Func<string?, double, string?, OfficeFontStyle, double> measure, string? defaultFontFamily = null) {
            if (!textBox.TryGetExportBoundsPoints(out double left, out double top, out double sourceWidth, out double sourceHeight))
                throw new ArgumentException("The text box has no renderable bounds.", nameof(textBox));
            var drawing = new OfficeDrawing(width, height);
            double scaleX = width / sourceWidth;
            double scaleY = height / sourceHeight;
            var mapping = new PowerPointShapeBoundsMapping(-left * scaleX, -top * scaleY, scaleX, scaleY);
            var colors = textBox.OwnerSlide == null ? null : GetSlideColorScheme(textBox.OwnerSlide);
            AddTextBox(drawing, textBox, diagnostics, mapping, colors, suppressFrame: true, measure: measure, defaultFontFamily: defaultFontFamily);
            foreach (OfficeDrawingElement element in drawing.Elements) {
                if (diagnostics.Any(diagnostic => diagnostic.Code == "POWERPOINT_TEXT_OVERFLOW")) break;
                ReportTextBoxOverflow(element, textBox, diagnostics, measure);
            }
            return drawing;
        }

        private static void ReportTextBoxOverflow(OfficeDrawingElement element, PowerPointTextBox textBox,
            List<OfficeImageExportDiagnostic> diagnostics, Func<string?, double, string?, OfficeFontStyle, double> measure) {
            bool clipped = false;
            if (element is OfficeDrawingRichText rich) {
                double availableHeight = Math.Max(0D, rich.Height - rich.Padding.Vertical);
                var layout = OfficeDrawingTextLayout.Create(rich, rich.Width - rich.Padding.Horizontal,
                    availableHeight, measure);
                clipped = layout.Clipped || OfficeTextPlacement.ResolveTop(0D, availableHeight, layout.Height, rich.VerticalAlignment) +
                    Math.Max(layout.Height, OfficeDrawingTextLayout.PaintedHeight(layout)) > availableHeight + 0.001D;
            } else if (element is OfficeDrawingText plain) {
                double size = plain.Font.Size;
                double lineHeightFactor = OfficeDrawingTextLayout.ResolveLineHeightFactor(plain.LineHeight, size);
                double availableHeight = Math.Max(0D, plain.Height - plain.Padding.Vertical);
                var layout = OfficeTextLayoutEngine.LayoutTextBlock(plain.Text, size,
                    plain.Width - plain.Padding.Horizontal, plain.Height - plain.Padding.Vertical,
                    lineHeightFactor, Math.Min(6D, size),
                    (value, fontSize) => measure(value, fontSize, plain.Font.FamilyName, plain.Font.Style),
                    plain.WrapText, shrinkToFit: plain.ShrinkToFit, paragraphIndent: plain.ParagraphIndent);
                clipped = layout.Clipped || OfficeTextPlacement.ResolveTop(0D, availableHeight, layout.Height, plain.VerticalAlignment) +
                    Math.Max(layout.Height, OfficeDrawingTextLayout.PaintedHeight(layout)) > availableHeight + 0.001D;
            }
            if (clipped) diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning, "POWERPOINT_TEXT_OVERFLOW",
                "Clipped PowerPoint text because it exceeds the available text frame.",
                DescribeShape(textBox), OfficeConversionLossKind.Approximation));
        }

        private static string ResolveTextBoxFallbackFont(PowerPointTextBox textBox, string? defaultFontFamily) =>
            defaultFontFamily ?? PowerPointTextDefaults.ResolveBodyLatinFont(textBox.OwnerSlide) ?? "Calibri";

        private static IEnumerable<OfficeRichTextRun> CreateEffectiveTextRuns(string text, PowerPointTextRun? run,
            PowerPointTextBox textBox, PowerPointParagraph paragraph, DocumentFormat.OpenXml.Drawing.ColorScheme? colors,
            PowerPointShapeBoundsMapping mapping, string? defaultFontFamily = null) {
            if (run == null) {
                yield return CreateRichTextRun(text, run, textBox, paragraph, colors, mapping, defaultFontFamily: defaultFontFamily);
                yield break;
            }
            foreach (PowerPointEffectiveTextSegment part in PowerPointEffectiveRunStyleResolver.ResolveSegments(
                run, paragraph, textBox.TextBody?.ListStyle, textBox.MasterTextStyle)) {
                PowerPointEffectiveRunStyle style = part.Style;
                string value = style.Capitalization is PowerPointCapitalization.AllCaps or PowerPointCapitalization.SmallCaps
                    ? OfficeTextCaseTransformer.Apply(part.Text, OfficeTextCase.Uppercase, ResolvePowerPointRunCulture(style.Language)) : part.Text;
                OfficeColor? authoredColor = OfficeOpenXmlThemeColorResolver.ResolveColor(run.RunProperties?.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>(), colors);
                OfficeColor color = authoredColor ?? (TryParseOfficeColor(style.Color, out OfficeColor effectiveColor)
                    ? effectiveColor : ResolveTextRunColor(run, textBox, colors));
                string? link = run.Hyperlink is { IsAbsoluteUri: true } uri ? uri.AbsoluteUri : null;
                var underline = MapUnderlineStyle(style.UnderlineStyle);
                if (link != null && underline == OfficeTextDecorationStyle.None) underline = OfficeTextDecorationStyle.Single;
                yield return new OfficeRichTextRun(value,
                    mapping.MapFontSize(style.FontSizePoints ?? textBox.FontSize ?? PowerPointTextDefaults.DefaultFontSizePoints),
                    color, style.Bold ?? textBox.Bold, style.Italic ?? textBox.Italic,
                    underline != OfficeTextDecorationStyle.None, style.FontName ?? textBox.FontName ?? ResolveTextBoxFallbackFont(textBox, defaultFontFamily),
                    style.StrikeStyle is { } strike && strike != PowerPointStrikeStyle.None,
                    ResolveTextRunBackgroundColor(run, colors), underline, MapStrikeStyle(style.StrikeStyle),
                    MapBaseline(style.BaselinePercent)) { LinkUri = link };
            }
        }
    }
}
