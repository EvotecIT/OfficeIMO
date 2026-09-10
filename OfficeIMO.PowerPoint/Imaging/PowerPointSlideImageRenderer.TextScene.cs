using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        internal static OfficeDrawing CreateTextBoxDrawing(PowerPointTextBox textBox,
            double width, double height, List<OfficeImageExportDiagnostic> diagnostics,
            Func<string?, double, string?, OfficeFontStyle, double> measure) {
            if (!textBox.TryGetExportBoundsPoints(out double left, out double top, out double sourceWidth, out double sourceHeight))
                throw new ArgumentException("The text box has no renderable bounds.", nameof(textBox));
            var drawing = new OfficeDrawing(width, height);
            var mapping = new PowerPointShapeBoundsMapping(-left, -top, 1D, 1D);
            var colors = textBox.OwnerSlide == null ? null : GetSlideColorScheme(textBox.OwnerSlide);
            AddTextBox(drawing, textBox, diagnostics, mapping, colors, suppressFrame: true, measure: measure);
            foreach (OfficeDrawingRichText text in drawing.Elements.OfType<OfficeDrawingRichText>()) {
                var layout = OfficeDrawingTextLayout.Create(text, text.Width - text.Padding.Horizontal,
                    text.Height - text.Padding.Vertical, measure);
                if (layout.Clipped) diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning, "POWERPOINT_TEXT_OVERFLOW",
                    "Clipped PowerPoint text because it exceeds the available text frame.",
                    DescribeShape(textBox), OfficeConversionLossKind.Approximation));
            }
            return drawing;
        }

        private static IEnumerable<OfficeRichTextRun> CreateEffectiveTextRuns(string text, PowerPointTextRun? run,
            PowerPointTextBox textBox, PowerPointParagraph paragraph, DocumentFormat.OpenXml.Drawing.ColorScheme? colors,
            PowerPointShapeBoundsMapping mapping) {
            if (run == null) {
                yield return CreateRichTextRun(text, run, textBox, paragraph, colors, mapping);
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
                    underline != OfficeTextDecorationStyle.None, style.FontName ?? textBox.FontName ?? PowerPointTextDefaults.ResolveBodyLatinFont(textBox.OwnerSlide) ?? "Calibri",
                    style.StrikeStyle is { } strike && strike != PowerPointStrikeStyle.None,
                    ResolveTextRunBackgroundColor(run, colors), underline, MapStrikeStyle(style.StrikeStyle),
                    MapBaseline(style.BaselinePercent)) { LinkUri = link };
            }
        }
    }
}
