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
            var mapping = new PowerPointShapeBoundsMapping(-left, -top, 1D, 1D);
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
                clipped = OfficeDrawingTextLayout.Create(rich, rich.Width - rich.Padding.Horizontal,
                    rich.Height - rich.Padding.Vertical, measure).Clipped;
            } else if (element is OfficeDrawingText plain) {
                double size = plain.Font.Size;
                double lineHeightFactor = plain.LineHeight.HasValue ? Math.Max(1D, plain.LineHeight.Value / size) : 1.2D;
                clipped = OfficeTextLayoutEngine.LayoutTextBlock(plain.Text, size,
                    plain.Width - plain.Padding.Horizontal, plain.Height - plain.Padding.Vertical,
                    lineHeightFactor, Math.Min(6D, size),
                    (value, fontSize) => measure(value, fontSize, plain.Font.FamilyName, plain.Font.Style),
                    plain.WrapText, shrinkToFit: plain.ShrinkToFit, paragraphIndent: plain.ParagraphIndent).Clipped;
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
