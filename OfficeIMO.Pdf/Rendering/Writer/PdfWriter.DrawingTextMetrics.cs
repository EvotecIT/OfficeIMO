using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Retains one drawing's font-measurement state only after its first text measurement.</summary>
    private sealed class DrawingTextMetricsState {
        private readonly PdfCreateOptions options;
        private readonly System.Threading.CancellationToken cancellationToken;
        private readonly DrawingTextMeasurementCanvases? canvases;
        private OfficeDrawingTextMetrics? value;

        internal DrawingTextMetricsState(PdfCreateOptions options, System.Threading.CancellationToken cancellationToken,
            DrawingTextMeasurementCanvases? canvases = null) {
            this.options = options;
            this.cancellationToken = cancellationToken;
            this.canvases = canvases;
        }

        internal OfficeDrawingTextMetrics Value => value ??= CreateDrawingTextMetrics(options, canvases, cancellationToken);
    }

    /// <summary>
    /// Measurement canvases for one layout pass. Parsing a font face is far more expensive than
    /// measuring with it, so every drawing in the pass reuses the canvas for the same face bytes,
    /// family, style and shaping configuration instead of rebuilding it per drawing. CFF execution
    /// budgets are not shared: each drawing's metrics charge their own allowance per face.
    /// </summary>
    private sealed class DrawingTextMeasurementCanvases {
        private const int MaximumCanvases = 256;
        private readonly Dictionary<(byte[] Data, string Family, OfficeFontStyle Style, IOfficeTextShapingProvider? Shaping, string? Language), OfficeRasterCanvas?> canvases =
            new Dictionary<(byte[], string, OfficeFontStyle, IOfficeTextShapingProvider?, string?), OfficeRasterCanvas?>();
        private readonly object gate = new object();

        internal OfficeRasterCanvas? GetOrCreate(byte[] data, string family, OfficeFontStyle style, PdfCreateOptions options,
            Dictionary<OfficeRasterCanvas, OfficeCffOperationBudget> budgets, System.Threading.CancellationToken cancellationToken) {
            var key = (data, family, style, options.TextShapingProvider, options.EffectiveTextShapingLanguage);
            lock (gate) {
                if (!canvases.TryGetValue(key, out OfficeRasterCanvas? canvas)) {
                    var fonts = new OfficeFontFaceCollection();
                    if (fonts.TryAdd(family, data, style)) {
                        canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), null, fonts,
                            options.TextShapingProvider, options.EffectiveTextShapingLanguage, cancellationToken: cancellationToken);
                    }
                    if (canvases.Count >= MaximumCanvases) canvases.Clear();
                    canvases[key] = canvas;
                }
                if (canvas != null) {
                    if (!budgets.TryGetValue(canvas, out OfficeCffOperationBudget? budget)) {
                        budget = new OfficeCffOperationBudget();
                        budgets.Add(canvas, budget);
                    }
                    canvas.UseCffOperationBudget(budget);
                }
                return canvas;
            }
        }
    }

    /// <summary>The face a metrics instance selected last, with the canvas and budget that selection installed.</summary>
    private sealed class DrawingTextCanvasSelection {
        private readonly byte[] data;
        private readonly string family;
        private readonly OfficeFontStyle style;
        private readonly IOfficeTextShapingProvider? shaping;
        private readonly string? language;

        internal DrawingTextCanvasSelection(byte[] data, string family, OfficeFontStyle style, IOfficeTextShapingProvider? shaping,
            string? language, OfficeRasterCanvas canvas, OfficeCffOperationBudget budget) {
            this.data = data; this.family = family; this.style = style; this.shaping = shaping; this.language = language;
            Canvas = canvas; Budget = budget;
        }

        internal OfficeRasterCanvas Canvas { get; }
        internal OfficeCffOperationBudget Budget { get; }

        internal bool Matches(byte[] data, string family, OfficeFontStyle style, IOfficeTextShapingProvider? shaping, string? language) =>
            ReferenceEquals(data, this.data) && style == this.style && ReferenceEquals(shaping, this.shaping)
            && string.Equals(family, this.family, StringComparison.Ordinal) && string.Equals(language, this.language, StringComparison.Ordinal);
    }

    // Font selection stays in PDF options; native geometry is measured by the shared font engine.
    internal static Func<string?, double, string?, OfficeFontStyle, double> CreateDrawingTextMeasure(PdfCreateOptions options) {
        return CreateDrawingTextMetrics(options).MeasureText;
    }

    internal static OfficeDrawingTextMetrics CreateDrawingTextMetrics(PdfCreateOptions options, System.Threading.CancellationToken cancellationToken = default) =>
        CreateDrawingTextMetrics(options, null, cancellationToken);

    private static OfficeDrawingTextMetrics CreateDrawingTextMetrics(PdfCreateOptions options,
        DrawingTextMeasurementCanvases? canvases, System.Threading.CancellationToken cancellationToken) {
        canvases ??= new DrawingTextMeasurementCanvases();
        // One allowance per face for this metrics instance, matching a canvas built for it alone.
        var budgets = new Dictionary<OfficeRasterCanvas, OfficeCffOperationBudget>();
        // Consecutive measurements nearly always select the same face. Reuse its canvas and this
        // instance's budget without the shared cache's lock and lookup; the face's identity is the
        // same byte array, family, style and shaping configuration GetOrCreate keys on.
        DrawingTextCanvasSelection? lastSelection = null;
        return new OfficeDrawingTextMetrics(Measure, MeasurePaint, MeasureHorizontalPaint, MeasurePositioned, MeasureSegmentPaint, MeasurePublisherFont, MeasureFeatures,
            MeasureSegmentHorizontalPaint, Prepare);

        IEnumerable<(PdfTextRun Run, OfficeRasterCanvas? Canvas, string Family, PdfStandardFont Font, bool SyntheticOblique, bool SyntheticBold)> Resolve(
            string? text, double size, string? family, OfficeFontStyle style, OfficeTextFeatureSettings? features = null) {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrEmpty(text)) yield break;
            PdfStandardFont font = PdfStandardFontMapper.TryMapFontFamily(family, out PdfStandardFont mapped)
                ? mapped : string.IsNullOrWhiteSpace(family) ? options.DefaultFont : PdfStandardFont.Helvetica;
            var run = new PdfTextRun(text!, bold: (style & OfficeFontStyle.Bold) != 0,
                italic: (style & OfficeFontStyle.Italic) != 0, fontSize: global::OfficeIMO.OfficeLength.Points(size), standardFont: font, fontFamily: family);
            if (features != null) run = run.WithFeatureSettings(features);
            foreach (PdfTextRun part in NormalizeFallbackRuns(new[] { run }, ChooseNormal(options.DefaultFont), options)) {
                cancellationToken.ThrowIfCancellationRequested();
                byte[]? data = null;
                string measuredFamily = part.FontFamily ?? family ?? "PDF default";
                PdfStandardFont selectedFont = ResolveFontForRun(part, ChooseNormal(options.DefaultFont));
                bool syntheticOblique = options.NeedsSyntheticOblique(selectedFont);
                bool syntheticBold = options.NeedsSyntheticBold(selectedFont);
                if (options.TryResolveNamedFontFace(part.FontFamily, part.Bold, part.Italic, out PdfNamedFontFace face)) {
                    options.TryGetNamedFontData(face, out data, out _);
                    measuredFamily = face.FamilyName;
                    syntheticOblique = options.NeedsSyntheticOblique(selectedFont, face);
                    syntheticBold = options.NeedsSyntheticBold(selectedFont, face);
                } else if (options.TryGetEmbeddedStandardFont(selectedFont, out PdfEmbeddedFont? embedded)) {
                    data = embedded!.DataSnapshot;
                }
                OfficeRasterCanvas? canvas = data != null ? GetCanvas(data, measuredFamily, style) : null;
                yield return (part, canvas, measuredFamily, selectedFont, syntheticOblique, syntheticBold);
            }
        }

        OfficeRasterCanvas? GetCanvas(byte[] data, string family, OfficeFontStyle style) {
            IOfficeTextShapingProvider? shaping = options.TextShapingProvider;
            string? language = options.EffectiveTextShapingLanguage;
            DrawingTextCanvasSelection? last = lastSelection;
            if (last != null && last.Matches(data, family, style, shaping, language)) {
                last.Canvas.UseCffOperationBudget(last.Budget);
                return last.Canvas;
            }
            OfficeRasterCanvas? canvas = canvases.GetOrCreate(data, family, style, options, budgets, cancellationToken);
            if (canvas != null) lastSelection = new DrawingTextCanvasSelection(data, family, style, shaping, language, canvas, budgets[canvas]);
            return canvas;
        }

        OfficePreparedTextShaping? Prepare(string text, double size, string family, OfficeFontStyle style, OfficeTextFeatureSettings? features) {
            var parts = Resolve(text, size, family, style, features).ToList();
            // Synthetic faces need the paint expansion applied by the per-run measurement path.
            if (parts.Count != 1 || parts[0].Canvas == null || parts[0].SyntheticOblique || parts[0].SyntheticBold) return null;
            var part = parts[0];
            // A source slice must select this same face again during PDF painting.
            // Composite fallback requests remain on their ordinary per-run route.
            if (!string.Equals(part.Family, family, StringComparison.OrdinalIgnoreCase)) return null;
            var prepared = part.Canvas!.PrepareRichText(text, size, part.Family, style, features);
            var selected = CreatePositionedTextSegment(part.Run, size, options);
            var font = ResolvePreparedDrawingFont(selected, options);
            return prepared != null && font != null && PdfExternalTextShaper.MatchesPreparedFont(prepared, font) ? prepared : null;
        }

        (double Left, double Right) MeasureSegmentHorizontalPaint(OfficeRichTextSegment segment, double size) {
            if (segment.PreparedShaping == null)
                return MeasureHorizontalPaintWithFeatures(segment.Text, size, segment.FontFamily, segment.FontStyle, segment.KerningFeatures);
            var part = Resolve(segment.Text, size, segment.FontFamily, segment.FontStyle, segment.KerningFeatures).Single();
            using var scope = part.Canvas!.PushPreparedShaping(segment.PreparedShaping);
            return part.Canvas.MeasureTextLineHorizontalPaintBoundsWithFeatures(segment.Text, size, part.Family,
                segment.FontStyle, segment.KerningFeatures);
        }

        OfficePublisherFontMetrics? MeasurePublisherFont(string? text, double size, string? family, OfficeFontStyle style) {
            var combined = new OfficePublisherFontMetrics.Union();
            foreach (var part in Resolve(string.IsNullOrEmpty(text) ? " " : text, size, family, style)) {
                double partSize = part.Run.FontSize?.ToPoints() ?? size;
                combined.Include(part.Canvas?.MeasurePublisherFont(part.Run.Text, partSize, part.Family, style), partSize);
            }
            return combined.Value;
        }

        double Measure(string? text, double size, string? family, OfficeFontStyle style) {
            double width = 0D;
            foreach (var part in Resolve(text, size, family, style))
                width += part.Canvas != null
                    ? part.Canvas.MeasureText(part.Run.Text, (part.Run.FontSize?.ToPoints()) ?? size, part.Family, style)
                    : GetRichSegmentWidth(CreatePositionedTextSegment(part.Run, size, options));
            return width;
        }

        double MeasureFeatures(string? text, double size, string? family, OfficeFontStyle style, OfficeTextFeatureSettings? features) {
            if (features == null) return Measure(text, size, family, style);
            double width = 0D;
            foreach (var part in Resolve(text, size, family, style, features))
                width += part.Canvas != null
                    ? part.Canvas.MeasureTextWithFeatures(part.Run.Text, part.Run.FontSize?.ToPoints() ?? size, part.Family, style, part.Run.FeatureSettings)
                    : GetRichSegmentWidth(CreatePositionedTextSegment(part.Run, size, options));
            return width;
        }

        double MeasurePositioned(string? text, double size, string? family, OfficeFontStyle style,
            OfficeTextDirection direction) {
            double width = 0D;
            foreach (var part in Resolve(text, size, family, style)) {
                width += part.Canvas != null
                    ? part.Canvas.MeasurePositionedText(
                        part.Run.Text, (part.Run.FontSize?.ToPoints()) ?? size, part.Family, style,
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
                double renderedSize = part.Run.FontSize?.ToPoints() ?? size;
                OfficeTextPaintBounds bounds = part.Canvas != null
                    ? part.Canvas.MeasureTextPaintBoundsWithFeatures(part.Run.Text, renderedSize, part.Family, style, features)
                    : GetStandardFontPaintBounds(part.Run.Text, part.Font, renderedSize);
                if (part.SyntheticBold) {
                    double pad = ContentStreamBuilder.SyntheticBoldStrokeWidth(renderedSize) / 2D;
                    bounds = new OfficeTextPaintBounds(bounds.Top - pad, bounds.Bottom + pad);
                }
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
                OfficeTextPaintBounds glyphs;
                if (segment.PreparedShaping != null) {
                    var part = Resolve(segment.Text, size, segment.FontFamily, segment.FontStyle, segment.KerningFeatures).Single();
                    using var scope = part.Canvas!.PushPreparedShaping(segment.PreparedShaping);
                    glyphs = part.Canvas.MeasureTextPaintBoundsWithFeatures(segment.Text, size, part.Family, segment.FontStyle, segment.KerningFeatures);
                } else glyphs = MeasurePaintWithFeatures(segment.Text, size, segment.FontFamily, segment.FontStyle, segment.KerningFeatures);
                Include(glyphs.Top, glyphs.Bottom);
            }
            if (segment.BackgroundColor.HasValue && segment.BackgroundColor.Value.A > 0 && segment.Width > 0D) {
                foreach (var part in Resolve(segment.Text, size, segment.FontFamily, segment.FontStyle)) {
                    var selected = CreatePositionedTextSegment(part.Run, size, options);
                    double fontSize = (part.Run.FontSize?.ToPoints()) ?? size;
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
                double renderedSize = part.Run.FontSize?.ToPoints() ?? size;
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
                if (part.SyntheticBold && !string.IsNullOrWhiteSpace(part.Run.Text)) {
                    double pad = ContentStreamBuilder.SyntheticBoldStrokeWidth(renderedSize) / 2D;
                    bounds.Left -= pad; bounds.Right += pad;
                }
                left = Math.Min(left, cursor + bounds.Left);
                right = Math.Max(right, cursor + bounds.Right);
                cursor += advance;
            }
            return (left, Math.Max(cursor, right));
        }
    }
}
