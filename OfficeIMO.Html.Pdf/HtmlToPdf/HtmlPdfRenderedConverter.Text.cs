using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static void AddText(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderText visual,
        RegisteredWebFonts webFonts,
        PdfCore.PdfConversionReport conversionReport,
        double surfaceWidth,
        bool asSpan,
        bool logicalTextOwned,
        CancellationToken cancellationToken,
        double? baselineFontSize = null,
        bool colorOpacityApplied = false,
        bool suppressLink = false,
        bool preservePositionedFrame = false,
        bool constrainToSurface = true,
        ClipBounds? logicalClip = null) {
        if (visual.Text.Length == 0) return;
        if (visual.WrappedLines != null) {
            foreach (HtmlRenderText fragment in visual.GetWrappedPaintFragments()) {
                cancellationToken.ThrowIfCancellationRequested();
                AddText(canvas, fragment, webFonts, conversionReport, surfaceWidth,
                    asSpan, logicalTextOwned, cancellationToken, baselineFontSize,
                    colorOpacityApplied, suppressLink, preservePositionedFrame, constrainToSurface, logicalClip);
            }
            return;
        }
        // Canvas and outline writers must anchor a script to the original line's metrics.
        baselineFontSize ??= visual.Font.Size;
        visual = visual.CopyStackingContextTo(visual.ResolveBaselineForPainting());
        if (visual.Y < 0D) {
            HtmlRenderText shifted = (HtmlRenderText)visual.Translate(0D, -visual.Y, visual.PaintOrder);
            canvas.Effect(OfficeTransform.Translate(0D, visual.Y * PointsPerCssPixel), 1D,
                nested => AddText(nested, shifted, webFonts, conversionReport, surfaceWidth, asSpan, logicalTextOwned, cancellationToken, baselineFontSize, suppressLink: suppressLink, preservePositionedFrame: preservePositionedFrame, constrainToSurface: constrainToSurface, logicalClip: logicalClip?.Translate(0D, -visual.Y)));
            return;
        }
        string? link = suppressLink || string.IsNullOrWhiteSpace(visual.Text) || IsFragmentLink(visual.LinkUri) ? null : visual.LinkUri;
        string? linkDestination = !suppressLink && IsFragmentLink(visual.LinkUri)
            ? MapNamedDestination(visual.LinkUri!.Substring(1))
            : null;
        double frameWidth = visual.Width;
        if (!preservePositionedFrame && visual.TextAdvanceWidth.HasValue) {
            double metricTolerance = Math.Max(
                visual.Font.Size,
                visual.TextAdvanceWidth.Value * 0.25D);
            frameWidth = Math.Max(frameWidth, visual.TextAdvanceWidth.Value + metricTolerance);
        }
        // An effect uses local coordinates. The destination page and authored
        // clips constrain its paint after transformation, not its local frame.
        frameWidth = Math.Max(0.01D, constrainToSurface
            ? Math.Min(frameWidth, Math.Max(0.01D, surfaceWidth - visual.X))
            : frameWidth);
        try {
            if (TryAddOutlinedText(
                    canvas,
                    visual,
                    webFonts,
                    conversionReport,
                    frameWidth,
                    asSpan,
                    logicalTextOwned,
                    cancellationToken,
                    baselineFontSize.Value,
                    suppressLink,
                    preservePositionedFrame,
                    logicalClip)) {
                return;
            }
        } catch (InvalidOperationException exception) when (
            exception.Message == "Font outline expansion exceeded the configured point budget."
            || exception.Message == "HTML-to-PDF outlined text exceeded the configured path-command budget.") {
            webFonts.OutlineBudget.StopOutlining();
            if (!webFonts.OutlineBudgetApproximationReported) {
                webFonts.OutlineBudgetApproximationReported = true;
                conversionReport.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.Html.Pdf",
                    HtmlPdfDiagnosticCodes.FontOutlineBudgetApproximated,
                    visual.Source ?? "html-text",
                    "The bounded font-outline budget was exhausted; remaining text used PDF text with possible font and shaping differences.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Approximation,
                    details: new Dictionary<string, string> {
                        ["FontFamily"] = visual.Font.FamilyName ?? string.Empty,
                        ["Fallback"] = "pdf-text"
                    }));
            }
        }
        if (!colorOpacityApplied && visual.Color.A < 255) {
            canvas.Effect(OfficeTransform.Identity, visual.Color.A / 255D,
                nested => AddText(nested, visual, webFonts, conversionReport, surfaceWidth,
                    asSpan, logicalTextOwned, cancellationToken, baselineFontSize, colorOpacityApplied: true, suppressLink: suppressLink, preservePositionedFrame: preservePositionedFrame, constrainToSurface: constrainToSurface, logicalClip: logicalClip));
            return;
        }
        OfficeFontStyle requestedStyle = (visual.Font.IsBold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (visual.Font.IsItalic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        string logicalPdfText = OmitUnavailablePrivateUseGlyphs(visual, webFonts, conversionReport, requestedStyle, cancellationToken);
        if (logicalPdfText.Length == 0) return;
        PdfCore.PdfTextRun[] CreatePdfRuns(string runText) {
            RunPlanPart[] parts = PlanPdfRuns(runText, visual, requestedStyle, webFonts);
            var runs = new PdfCore.PdfTextRun[parts.Length];
            for (int index = 0; index < parts.Length; index++) {
                RunPlanPart part = parts[index];
                runs[index] = new PdfCore.PdfTextRun(
                    part.Text,
                    bold: part.Bold,
                    underline: visual.Font.IsUnderline,
                    color: (visual.Color.A == 0 ? (global::OfficeIMO.OfficeColor?)null : visual.Color),
                    italic: part.Italic,
                    strike: visual.Font.IsStrikethrough,
                    fontSize: global::OfficeIMO.OfficeLength.Points(visual.Font.Size * PointsPerCssPixel),
                    standardFont: part.StandardFont,
                    linkUri: link,
                    linkContents: link == null ? null : runText,
                    linkDestinationName: linkDestination,
                    fontFamily: part.Family,
                    baseline: MapTextBaseline(visual.Baseline),
                    underlineStyle: visual.UnderlineStyle,
                    strikeStyle: visual.StrikethroughStyle,
                    decorationColor: (visual.DecorationColor.A == 0 ? (global::OfficeIMO.OfficeColor?)null : visual.DecorationColor))
                    .WithFeatureSettings(visual.FeatureSettings)
                    .WithPaintedGlyphPreservation(visual.PreservesPaintedGlyphs);
            }
            return runs;
        }
        string pdfText = ReplaceUnpaintableTextElements(logicalPdfText, visual, webFonts, CreatePdfRuns, cancellationToken);
        PdfCore.PdfTextRun[] runs = CreatePdfRuns(pdfText);
        void Paint(PdfCore.PdfPageCanvas target) => target.PositionedText(
            runs,
            asSpan ? PdfCore.PdfCanvasTextStructureRole.Span : MapStructureRole(visual.SemanticRole),
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            frameWidth * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            (visual.Color.A == 0 ? (global::OfficeIMO.OfficeColor?)null : visual.Color),
            MapAlignment(visual.Alignment),
            baselineFontSize.Value * PointsPerCssPixel,
            visual.PaintLineHeight * PointsPerCssPixel,
            (visual.TextPaintWidth ?? visual.TextAdvanceWidth) * PointsPerCssPixel,
            visual.PaintTopOverflow * PointsPerCssPixel,
            fontMetricScale: PointsPerCssPixel, logicalOrder: visual.SemanticSourceOrder);
        double logicalWidth = visual.TextPaintWidth ?? visual.TextAdvanceWidth ?? frameWidth;
        // The source-em paint carrier may start above the line box. Its blank
        // leading must not turn otherwise native visible glyphs into an
        // invisible logical replacement merely because it crosses a page edge.
        double logicalY = visual.UsesOrdinaryLineCarrier && visual.RelativePaintOffsetY == 0D
            ? Math.Max(visual.Y, visual.LayoutY) : visual.Y;
        double logicalHeight = Math.Max(0.01D, Math.Min(visual.Height, visual.Font.Size));
        var carrier = logicalClip?.ConstrainLogicalRectangle(visual.X, logicalY, logicalWidth, logicalHeight)
            ?? (X: visual.X, Y: logicalY, Width: logicalWidth, Height: logicalHeight);
        // Native text already carries transform-aware glyph geometry. Project
        // only page-space clipped native text; transformed logical/outlined
        // groups retain their own replacement carriers without changing the
        // native reader's existing partially-visible text policy.
        // A replacement glyph must not reach text extraction in place of the source character.
        bool replacedText = !ReferenceEquals(pdfText, logicalPdfText);
        if (!logicalTextOwned && (replacedText || (!logicalClip.HasValue || logicalClip.Value.ConstrainToSurface)
            && (carrier.X != visual.X || carrier.Y != logicalY
            || carrier.Width != logicalWidth || carrier.Height != logicalHeight))) {
            canvas.ActualText(logicalPdfText, carrier.X * PointsPerCssPixel, carrier.Y * PointsPerCssPixel,
                carrier.Width * PointsPerCssPixel, carrier.Height * PointsPerCssPixel, Paint, logicalOrder: visual.SemanticSourceOrder);
        } else Paint(canvas);
    }


    private const int MaximumRunPlanTextLength = 256;
    private const int MaximumRunPlans = 8192;

    /// <summary>
    /// Plans the font fallback and face style of a painted run. The plan depends only on the run
    /// text, requested family and face, and the operation's registered fonts, so a table or report
    /// that repeats a cell value plans it once per conversion.
    /// </summary>
    private static RunPlanPart[] PlanPdfRuns(string runText, HtmlRenderText visual, OfficeFontStyle requestedStyle,
        RegisteredWebFonts webFonts) {
        var key = new RunPlanKey(runText, visual.Font.FamilyName, requestedStyle, visual.FontDescriptor,
            webFonts.AllowInstalledFontFaces);
        if (webFonts.RunPlans.TryGetValue(key, out RunPlanPart[]? known)) return known;
        var parts = new List<RunPlanPart>();
        foreach (OfficeFontFallbackRun fallbackRun in webFonts.Faces.PlanFallbackRuns(runText, visual.Font.FamilyName, requestedStyle)) {
            bool allowInstalledFace = webFonts.AllowInstalledFontFaces
                && !webFonts.Slots.ContainsKey(fallbackRun.FamilyName);
            string family = ResolvePdfFontFamilyForText(
                fallbackRun.FamilyRequest,
                fallbackRun.Text,
                visual.Font.IsBold,
                visual.Font.IsItalic,
                visual.FontDescriptor,
                allowInstalledFace,
                webFonts.Options);
            foreach (NamedFaceStyleRun faceRun in PlanNamedFaceStyleRuns(
                         fallbackRun.Text, family, visual.Font.IsBold, visual.Font.IsItalic, webFonts.Options)) {
                parts.Add(new RunPlanPart(faceRun.Text, faceRun.Bold, faceRun.Italic, family,
                    MapFont(family, faceRun.Text, requestedStyle, webFonts)));
            }
        }
        RunPlanPart[] plan = parts.ToArray();
        if (runText.Length <= MaximumRunPlanTextLength && webFonts.RunPlans.Count < MaximumRunPlans) webFonts.RunPlans.Add(key, plan);
        return plan;
    }

    private readonly struct RunPlanPart {
        internal RunPlanPart(string text, bool bold, bool italic, string family, PdfCore.PdfStandardFont standardFont) {
            Text = text;
            Bold = bold;
            Italic = italic;
            Family = family;
            StandardFont = standardFont;
        }

        internal string Text { get; }
        internal bool Bold { get; }
        internal bool Italic { get; }
        internal string Family { get; }
        internal PdfCore.PdfStandardFont StandardFont { get; }
    }

    private readonly struct RunPlanKey : IEquatable<RunPlanKey> {
        private readonly string _text;
        private readonly string? _family;
        private readonly OfficeFontStyle _style;
        private readonly OfficeFontFaceDescriptor _descriptor;
        private readonly bool _allowInstalledFaces;

        internal RunPlanKey(string text, string? family, OfficeFontStyle style, OfficeFontFaceDescriptor descriptor, bool allowInstalledFaces) {
            _text = text;
            _family = family;
            _style = style;
            _descriptor = descriptor;
            _allowInstalledFaces = allowInstalledFaces;
        }

        public bool Equals(RunPlanKey other) =>
            _style == other._style && _allowInstalledFaces == other._allowInstalledFaces
            && _descriptor.Equals(other._descriptor)
            && string.Equals(_text, other._text, StringComparison.Ordinal)
            && string.Equals(_family, other._family, StringComparison.Ordinal);

        public override bool Equals(object? obj) => obj is RunPlanKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = StringComparer.Ordinal.GetHashCode(_text);
                hash = (hash * 397) ^ (_family == null ? 0 : StringComparer.Ordinal.GetHashCode(_family));
                hash = (hash * 397) ^ (int)_style;
                hash = (hash * 397) ^ _descriptor.GetHashCode();
                return (hash * 397) ^ (_allowInstalledFaces ? 1 : 0);
            }
        }
    }
}
