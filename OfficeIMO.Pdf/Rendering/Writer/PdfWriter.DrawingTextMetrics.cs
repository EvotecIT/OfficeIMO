using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Font selection stays in PDF options; native geometry is measured by the shared font engine.
    internal static Func<string?, double, string?, OfficeFontStyle, double> CreateDrawingTextMeasure(PdfOptions options) {
        return CreateDrawingTextMetrics(options).MeasureText;
    }

    internal static OfficeDrawingTextMetrics CreateDrawingTextMetrics(PdfOptions options) {
        var metrics = new Dictionary<(byte[] Data, string Family, OfficeFontStyle Style), OfficeRasterCanvas?>();
        return new OfficeDrawingTextMetrics(Measure, MeasurePaint);

        IEnumerable<(PdfTextRun Run, OfficeRasterCanvas? Canvas, string Family, PdfStandardFont Font)> Resolve(
            string? text, double size, string? family, OfficeFontStyle style) {
            if (string.IsNullOrEmpty(text)) yield break;
            PdfStandardFont font = PdfStandardFontMapper.TryMapFontFamily(family, out PdfStandardFont mapped)
                ? mapped : string.IsNullOrWhiteSpace(family) ? options.DefaultFont : PdfStandardFont.Helvetica;
            var run = new PdfTextRun(text!, bold: (style & OfficeFontStyle.Bold) != 0,
                italic: (style & OfficeFontStyle.Italic) != 0, fontSize: size, font: font, fontFamily: family);
            foreach (PdfTextRun part in NormalizeFallbackRuns(new[] { run }, ChooseNormal(options.DefaultFont), options)) {
                byte[]? data = null;
                string measuredFamily = part.FontFamily ?? family ?? "PDF default";
                PdfStandardFont selectedFont = ResolveFontForRun(part, ChooseNormal(options.DefaultFont));
                if (options.TryResolveNamedFontFace(part.FontFamily, part.Bold, part.Italic, out PdfNamedFontFace face)) {
                    options.TryGetNamedFontData(face, out data, out _);
                    measuredFamily = face.FamilyName;
                } else if (options.TryGetEmbeddedStandardFont(selectedFont, out PdfEmbeddedFont? embedded)) {
                    data = embedded!.DataSnapshot;
                }
                OfficeRasterCanvas? canvas = null;
                if (data != null) {
                    var key = (data, measuredFamily, style);
                    if (!metrics.TryGetValue(key, out canvas)) {
                        var fonts = new OfficeFontFaceCollection();
                        if (fonts.TryAdd(measuredFamily, data, style)) {
                            canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), null, fonts,
                                options.TextShapingProvider, options.Language);
                        }
                        if (metrics.Count >= 256) metrics.Clear();
                        metrics[key] = canvas;
                    }
                }
                yield return (part, canvas, measuredFamily, selectedFont);
            }
        }

        double Measure(string? text, double size, string? family, OfficeFontStyle style) {
            double width = 0D;
            foreach (var part in Resolve(text, size, family, style))
                width += part.Canvas != null
                    ? part.Canvas.MeasureText(part.Run.Text, part.Run.FontSize ?? size, part.Family, style)
                    : MeasureRichSegment(CreatePositionedTextSegment(part.Run, size, options), options);
            return width;
        }

        OfficeTextPaintBounds MeasurePaint(string? text, double size, string? family, OfficeFontStyle style) {
            double top = -size * .84D, bottom = size * .16D;
            foreach (var part in Resolve(text, size, family, style)) {
                OfficeTextPaintBounds bounds = part.Canvas != null
                    ? part.Canvas.MeasureTextPaintBounds(part.Run.Text, part.Run.FontSize ?? size, part.Family, style)
                    : GetStandardFontPaintBounds(part.Font, part.Run.FontSize ?? size);
                top = Math.Min(top, bounds.Top); bottom = Math.Max(bottom, bounds.Bottom);
            }
            return new OfficeTextPaintBounds(top, bottom);
        }
    }
}
