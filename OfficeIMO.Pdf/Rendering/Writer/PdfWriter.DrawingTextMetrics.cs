using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Font selection stays in PDF options; native geometry is measured by the shared font engine.
    internal static Func<string?, double, string?, OfficeFontStyle, double> CreateDrawingTextMeasure(PdfOptions options) {
        return CreateDrawingTextMetrics(options).MeasureText;
    }

    internal static OfficeDrawingTextMetrics CreateDrawingTextMetrics(PdfOptions options, System.Threading.CancellationToken cancellationToken = default) {
        var metrics = new Dictionary<(byte[] Data, string Family, OfficeFontStyle Style), OfficeRasterCanvas?>();
        return new OfficeDrawingTextMetrics(Measure, MeasurePaint, MeasureHorizontalPaint, MeasurePositioned, MeasureSegmentPaint, MeasurePublisherFont, MeasureFeatures,
            (segment, size) => MeasureHorizontalPaintWithFeatures(segment.Text, size, segment.FontFamily, segment.FontStyle, segment.KerningFeatures));

        IEnumerable<(PdfTextRun Run, OfficeRasterCanvas? Canvas, string Family, PdfStandardFont Font, bool SyntheticOblique)> Resolve(
            string? text, double size, string? family, OfficeFontStyle style, OfficeTextFeatureSettings? features = null) {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrEmpty(text)) yield break;
            PdfStandardFont font = PdfStandardFontMapper.TryMapFontFamily(family, out PdfStandardFont mapped)
                ? mapped : string.IsNullOrWhiteSpace(family) ? options.DefaultFont : PdfStandardFont.Helvetica;
            var run = new PdfTextRun(text!, bold: (style & OfficeFontStyle.Bold) != 0,
                italic: (style & OfficeFontStyle.Italic) != 0, fontSize: size, font: font, fontFamily: family);
            if (features != null) run = run.WithFeatureSettings(features);
            foreach (PdfTextRun part in NormalizeFallbackRuns(new[] { run }, ChooseNormal(options.DefaultFont), options)) {
                cancellationToken.ThrowIfCancellationRequested();
                byte[]? data = null;
                string measuredFamily = part.FontFamily ?? family ?? "PDF default";
                PdfStandardFont selectedFont = ResolveFontForRun(part, ChooseNormal(options.DefaultFont));
                bool syntheticOblique = options.NeedsSyntheticOblique(selectedFont);
                if (options.TryResolveNamedFontFace(part.FontFamily, part.Bold, part.Italic, out PdfNamedFontFace face)) {
                    options.TryGetNamedFontData(face, out data, out _);
                    measuredFamily = face.FamilyName;
                    syntheticOblique = options.NeedsSyntheticOblique(selectedFont, face);
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
                                options.TextShapingProvider, options.EffectiveTextShapingLanguage, cancellationToken: cancellationToken);
                        }
                        if (metrics.Count >= 256) metrics.Clear();
                        metrics[key] = canvas;
                    }
                }
                yield return (part, canvas, measuredFamily, selectedFont, syntheticOblique);
            }
        }

        OfficePublisherFontMetrics? MeasurePublisherFont(string? text, double size, string? family, OfficeFontStyle style) {
            var combined = new OfficePublisherFontMetrics.Union();
            foreach (var part in Resolve(string.IsNullOrEmpty(text) ? " " : text, size, family, style)) {
                double partSize = part.Run.FontSize ?? size;
                combined.Include(part.Canvas?.MeasurePublisherFont(part.Run.Text, partSize, part.Family, style), partSize);
            }
            return combined.Value;
        }

        double Measure(string? text, double size, string? family, OfficeFontStyle style) {
            double width = 0D;
            foreach (var part in Resolve(text, size, family, style))
                width += part.Canvas != null
                    ? part.Canvas.MeasureText(part.Run.Text, part.Run.FontSize ?? size, part.Family, style)
                    : GetRichSegmentWidth(CreatePositionedTextSegment(part.Run, size, options));
            return width;
        }

        double MeasureFeatures(string? text, double size, string? family, OfficeFontStyle style, OfficeTextFeatureSettings? features) {
            if (features == null) return Measure(text, size, family, style);
            double width = 0D;
            foreach (var part in Resolve(text, size, family, style, features))
                width += part.Canvas != null
                    ? part.Canvas.MeasureTextWithFeatures(part.Run.Text, part.Run.FontSize ?? size, part.Family, style, part.Run.FeatureSettings)
                    : GetRichSegmentWidth(CreatePositionedTextSegment(part.Run, size, options));
            return width;
        }

        double MeasurePositioned(string? text, double size, string? family, OfficeFontStyle style,
            OfficeTextDirection direction) {
            double width = 0D;
            foreach (var part in Resolve(text, size, family, style)) {
                width += part.Canvas != null
                    ? part.Canvas.MeasurePositionedText(
                        part.Run.Text, part.Run.FontSize ?? size, part.Family, style,
                        part.Run.FeatureSettings, direction)
                    : GetRichSegmentWidth(CreatePositionedTextSegment(
                        part.Run.WithTextDirection(direction), size, options));
            }
            return width;
        }

        OfficeTextPaintBounds MeasurePaint(string? text, double size, string? family, OfficeFontStyle style) =>
            MeasurePaintWithFeatures(text, size, family, style, null);

        OfficeTextPaintBounds MeasurePaintWithFeatures(string? text, double size, string? family, OfficeFontStyle style, OfficeTextFeatureSettings? features) {
            double top = 0D, bottom = 0D;
            bool hasPaint = false;
            foreach (var part in Resolve(text, size, family, style, features)) {
                if (string.IsNullOrWhiteSpace(part.Run.Text)) continue;
                OfficeTextPaintBounds bounds = part.Canvas != null
                    ? part.Canvas.MeasureTextPaintBoundsWithFeatures(part.Run.Text, part.Run.FontSize ?? size, part.Family, style, features)
                    : GetStandardFontPaintBounds(part.Run.Text, part.Font, part.Run.FontSize ?? size);
                if (!hasPaint) {
                    top = bounds.Top; bottom = bounds.Bottom; hasPaint = true;
                } else {
                    top = Math.Min(top, bounds.Top); bottom = Math.Max(bottom, bounds.Bottom);
                }
            }
            return new OfficeTextPaintBounds(top, bottom);
        }

        // Drawing segments share the writer's actual font selection, background
        // ascender/descender metrics and decoration contours. Glyph bounds alone
        // cannot establish that a tight frame retains its supported paint.
        OfficeTextPaintBounds? MeasureSegmentPaint(OfficeRichTextSegment segment, double size) {
            double top = double.PositiveInfinity, bottom = double.NegativeInfinity;
            if (segment.Color.A > 0 && !string.IsNullOrWhiteSpace(segment.Text)) {
                OfficeTextPaintBounds glyphs = MeasurePaintWithFeatures(segment.Text, size, segment.FontFamily, segment.FontStyle, segment.KerningFeatures);
                Include(glyphs.Top, glyphs.Bottom);
            }
            if (segment.BackgroundColor.HasValue && segment.BackgroundColor.Value.A > 0 && segment.Width > 0D) {
                foreach (var part in Resolve(segment.Text, size, segment.FontFamily, segment.FontStyle)) {
                    var selected = CreatePositionedTextSegment(part.Run, size, options);
                    double fontSize = part.Run.FontSize ?? size;
                    double padding = Math.Max(.45D, fontSize * .05D);
                    Include(-GetAscenderForOptions(selected.Font, selected.NamedFont, fontSize, options) - padding,
                        GetDescenderForOptions(selected.Font, selected.NamedFont, fontSize, options) + padding);
                }
            }
            if (segment.Color.A > 0 && segment.Width > 0D && segment.Text.Length > 0) {
                Decoration(segment.UnderlineStyle, -size * .15D);
                Decoration(segment.StrikethroughStyle, size * .32D);
            }
            return double.IsPositiveInfinity(top) ? null : new OfficeTextPaintBounds(top, bottom);

            void Include(double start, double end) { top = Math.Min(top, start); bottom = Math.Max(bottom, end); }
            void Decoration(OfficeTextDecorationStyle style, double y) {
                if (style == OfficeTextDecorationStyle.None || style == OfficeTextDecorationStyle.Words && string.IsNullOrWhiteSpace(segment.Text)) return;
                double pdfBottom = double.PositiveInfinity, pdfTop = double.NegativeInfinity;
                IncludeDecorationVerticalBounds(style, y, .5D, ref pdfBottom, ref pdfTop);
                Include(-pdfTop, -pdfBottom);
            }
        }

        (double Left, double Right) MeasureHorizontalPaint(string? text, double size, string? family, OfficeFontStyle style) =>
            MeasureHorizontalPaintWithFeatures(text, size, family, style, null);

        (double Left, double Right) MeasureHorizontalPaintWithFeatures(string? text, double size, string? family, OfficeFontStyle style, OfficeTextFeatureSettings? features) {
            double cursor = 0D, left = 0D, right = 0D;
            foreach (var part in Resolve(text, size, family, style, features)) {
                double renderedSize = part.Run.FontSize ?? size;
                double advance = part.Canvas != null
                    ? part.Canvas.MeasureTextWithFeatures(part.Run.Text, renderedSize, part.Family, style, features)
                    : GetRichSegmentWidth(CreatePositionedTextSegment(part.Run, size, options));
                var bounds = string.IsNullOrWhiteSpace(part.Run.Text) ? (Left: 0D, Right: advance) :
                    part.Canvas != null
                        ? part.Canvas.MeasureTextLineHorizontalPaintBoundsWithFeatures(part.Run.Text, renderedSize, part.Family, style, features)
                        : GetStandardFontHorizontalPaintBounds(part.Font, renderedSize, advance);
                if (part.SyntheticOblique && !string.IsNullOrWhiteSpace(part.Run.Text)) {
                    // The selected fallback bytes are unsheared. Bound the exact PDF
                    // text-matrix shear rather than the raster renderer's synthetic style.
                    OfficeTextPaintBounds vertical = part.Canvas != null
                        ? part.Canvas.MeasureTextPaintBoundsWithFeatures(part.Run.Text, renderedSize, part.Family, style, features)
                        : GetStandardFontPaintBounds(part.Run.Text, part.Font, renderedSize);
                    bounds.Left += Math.Min(-vertical.Top, -vertical.Bottom) * ContentStreamBuilder.SyntheticObliqueShear;
                    bounds.Right += Math.Max(-vertical.Top, -vertical.Bottom) * ContentStreamBuilder.SyntheticObliqueShear;
                }
                left = Math.Min(left, cursor + bounds.Left);
                right = Math.Max(right, cursor + bounds.Right);
                cursor += advance;
            }
            return (left, Math.Max(cursor, right));
        }
    }
}
