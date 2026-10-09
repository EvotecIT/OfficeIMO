using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal sealed partial class VisioDrawingProjection {
    private void ProjectText(OfficeDrawing drawing, string text, VisioTextStyle? style, VisioRichTextProjection? rich,
        double x, double y, double width, double height, double angle, double defaultSize, string location) {
        if (width <= 0 || height <= 0 || double.IsNaN(width) || double.IsNaN(height) || double.IsInfinity(width) || double.IsInfinity(height)) {
            _report.Add("VISIO_DRAWING_TEXT_FRAME", "Text is omitted because its cached content frame is empty or invalid.", OfficeConversionLossKind.Omission, location);
            return;
        }
        if (text.Length > OfficeTextLayoutEngine.MaximumLayoutTextCharacters) {
            _report.Add("VISIO_DRAWING_TEXT_LIMIT", "Text exceeds the shared drawing layout character limit.", OfficeConversionLossKind.Omission, location);
            return;
        }
        var local = new OfficeDrawing(width, height);
        local.Fonts.AddRange(drawing.Fonts);
        local.TextShapingProvider = drawing.TextShapingProvider;
        local.TextShapingLanguage = drawing.TextShapingLanguage;
        if (style?.BackgroundColor is OfficeColor background && background.A > 0) {
            double opacity = 1D - Math.Max(0, Math.Min(100, style.BackgroundTransparency ?? 0)) / 100D;
            OfficeShape backing = OfficeShape.Rectangle(width, height);
            backing.FillColor = OfficeColor.FromRgba(background.R, background.G, background.B, (byte)Math.Round(background.A * opacity));
            local.AddShape(backing, 0, 0);
        }
        OfficeTextVerticalAlignment vertical = VisioDrawingTextAlignment.ToOfficeTextVerticalAlignment(style?.VerticalAlignment);
        if (rich?.Paragraphs.Count > 0) local.AddRichTextParagraphsCore(rich.Paragraphs, 0, 0, width, height, vertical, true, null, shrinkToFit: true);
        else if (rich != null) local.AddRichText(rich.Runs, 0, 0, width, height, rich.Alignment, verticalAlignment: vertical, shrinkToFit: true);
        else {
            string display = style?.SmallCaps == true || style?.Capitalization == VisioTextCapitalization.AllCaps
                ? OfficeTextCaseTransformer.Apply(text, OfficeTextCase.Uppercase, CultureInfo.InvariantCulture)
                : style?.Capitalization == VisioTextCapitalization.InitialCaps
                    ? OfficeTextCaseTransformer.Apply(text, OfficeTextCase.Capitalize, CultureInfo.InvariantCulture) : text;
            var run = new OfficeRichTextRun(display, style?.Size ?? defaultSize, style?.Color ?? OfficeColor.FromRgb(17, 24, 39),
                style?.Bold == true, style?.Italic == true, style?.Underline == true, style?.FontFamily ?? "Arial",
                style?.Strikethrough == true, underlineStyle: style?.UnderlineStyle ?? OfficeTextDecorationStyle.None,
                strikethroughStyle: style?.StrikethroughStyle ?? OfficeTextDecorationStyle.None,
                baseline: style?.Baseline ?? OfficeTextBaseline.Normal);
            local.AddRichTextParagraphsCore(new[] { new OfficeRichTextParagraph(new[] { run },
                VisioDrawingTextAlignment.ToOfficeTextAlignment(style?.HorizontalAlignment)) }, 0, 0, width, height, vertical, true, null, shrinkToFit: true);
        }
        OfficeDrawingRichText positioned = local.Elements.OfType<OfficeDrawingRichText>().Single();
        OfficeDrawingTextMetrics? metrics = _options.LayoutMetrics;
        OfficeRichTextBlockLayout layout;
        if (metrics != null) layout = OfficeDrawingTextLayout.Create(positioned, width, height,
            metrics.MeasureText, measurePaint: metrics.MeasurePaintBounds);
        else {
            OfficeRasterCanvas rasterMetrics = OfficeDrawingTextLayout.CreateMetrics(local, _token);
            layout = OfficeDrawingTextLayout.Create(positioned, width, height,
                rasterMetrics.MeasureText, measurePaint: rasterMetrics.MeasureTextPaintBounds);
        }
        if (layout.Clipped)
            _report.Add("VISIO_DRAWING_TEXT_CLIPPED", "The shared fitted text layout cannot retain all content inside the cached frame with these font resources.", OfficeConversionLossKind.Omission, location);
        drawing.AddEffectDrawing(local, OfficeTransform.Translate(-width / 2D, -height / 2D)
            .Then(OfficeTransform.RotateDegrees(-angle * 180D / Math.PI)).Then(OfficeTransform.Translate(x, y)));
        _report.Add("VISIO_DRAWING_TEXT_LAYOUT", "Supported cached text runs and paragraphs use shared font fitting and frame clipping; native fitting, exact font placement, and advanced typography are not qualified.", OfficeConversionLossKind.Approximation, location);
    }
}
