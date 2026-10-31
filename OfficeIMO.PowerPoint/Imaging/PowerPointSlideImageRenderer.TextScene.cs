using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        internal static OfficeDrawing CreateTextBoxDrawing(PowerPointTextBox textBox,
            double width, double height, List<OfficeImageExportDiagnostic> diagnostics,
            Func<string?, double, string?, OfficeFontStyle, double> measure, string? defaultFontFamily = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!textBox.TryGetExportBoundsPoints(out double left, out double top, out double sourceWidth, out double sourceHeight))
                throw new ArgumentException("The text box has no renderable bounds.", nameof(textBox));
            double scaleX = width / sourceWidth;
            double scaleY = height / sourceHeight;
            var mapping = new PowerPointShapeBoundsMapping(-left * scaleX, -top * scaleY, scaleX, scaleY);
            var colors = textBox.OwnerSlide == null ? null : GetSlideColorScheme(textBox.OwnerSlide);
            return CreateFittedTextBoxDrawing(textBox, width, height, diagnostics, mapping, colors, measure, defaultFontFamily, cancellationToken);
        }

        private static string ResolveTextBoxFallbackFont(PowerPointTextBox textBox, string? defaultFontFamily) =>
            defaultFontFamily ?? PowerPointTextDefaults.ResolveBodyLatinFont(textBox.OwnerSlide) ?? "Calibri";

        private static void ReportCharacterSpacing(PowerPointTextBox textBox, List<OfficeImageExportDiagnostic> diagnostics,
            CancellationToken cancellationToken) {
            var inheritedSpacingByLevel = new Dictionary<int, int?>();
            if (!textBox.Paragraphs.Any(paragraph => paragraph.InlineNodes.Any(node => {
                cancellationToken.ThrowIfCancellationRequested();
                return node.Run is { } run && !string.IsNullOrEmpty(run.Text)
                    && PowerPointEffectiveRunStyleResolver.ResolveCharacterSpacingPoints(run, paragraph, textBox,
                        inheritedSpacingByLevel).GetValueOrDefault() != 0;
            }))) return;
            diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                "powerpoint-character-spacing-approximated",
                "Shared text rendering uses default character spacing; authored spacing is retained in the presentation.",
                DescribeShape(textBox), OfficeConversionLossKind.Approximation));
        }

        private static IEnumerable<OfficeRichTextRun> CreateEffectiveTextRuns(string text, PowerPointTextRun? run,
            PowerPointTextBox textBox, PowerPointParagraph paragraph, DocumentFormat.OpenXml.Drawing.ColorScheme? colors,
            PowerPointShapeBoundsMapping mapping, string? defaultFontFamily = null, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            if (run == null) {
                yield return CreateRichTextRun(text, run, textBox, paragraph, colors, mapping, defaultFontFamily: defaultFontFamily);
                yield break;
            }
            foreach (PowerPointEffectiveTextSegment part in PowerPointEffectiveRunStyleResolver.ResolveSegments(
                run, paragraph, textBox.TextBody?.ListStyle, textBox.MasterTextStyle)) {
                cancellationToken.ThrowIfCancellationRequested();
                PowerPointEffectiveRunStyle style = part.Style;
                string value = style.Capitalization is PowerPointCapitalization.AllCaps or PowerPointCapitalization.SmallCaps
                    ? OfficeTextCaseTransformer.Apply(part.Text, OfficeTextCase.Uppercase, ResolvePowerPointRunCulture(style.Language)) : part.Text;
                OfficeColor color = PowerPointEffectiveRunStyleResolver.ResolveTextColorForExport(run, paragraph, textBox)
                    ?? OfficeColor.Black;
                string? link = run.Hyperlink is { IsAbsoluteUri: true } uri ? uri.AbsoluteUri : null;
                string? destination = PowerPointHyperlinkResolver.TryParseSlideFragment(run.Hyperlink, out int targetSlide)
                    ? PowerPointHyperlinkResolver.GetSlideDestinationName(targetSlide) : null;
                var underline = MapUnderlineStyle(style.UnderlineStyle);
                if (link != null && underline == OfficeTextDecorationStyle.None) underline = OfficeTextDecorationStyle.Single;
                yield return new OfficeRichTextRun(value,
                    mapping.MapFontSize(style.FontSizePoints ?? textBox.FontSize?.ToPoints() ?? PowerPointTextDefaults.DefaultFontSizePoints),
                    color, style.Bold ?? textBox.Bold, style.Italic ?? textBox.Italic,
                    underline != OfficeTextDecorationStyle.None, style.FontName ?? textBox.FontName ?? ResolveTextBoxFallbackFont(textBox, defaultFontFamily),
                    style.StrikeStyle is { } strike && strike != PowerPointStrikeStyle.None,
                    ResolveTextRunBackgroundColor(run, textBox, colors), underline, MapStrikeStyle(style.StrikeStyle),
                    MapBaseline(style.BaselinePercent)) { LinkUri = link, LinkDestinationName = destination };
            }
        }
    }
}
