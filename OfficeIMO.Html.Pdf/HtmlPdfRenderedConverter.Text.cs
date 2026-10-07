using OfficeIMO.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
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
        visual = visual.ResolveBaselineForPainting();
        if (visual.Y < 0D) {
            HtmlRenderText shifted = (HtmlRenderText)visual.TranslatePaint(0D, -visual.Y, visual.PaintOrder);
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
        string pdfText = OmitUnavailablePrivateUseGlyphs(visual, webFonts, conversionReport, requestedStyle, cancellationToken);
        if (pdfText.Length == 0) return;
        var runs = webFonts.Faces.PlanFallbackRuns(pdfText, visual.Font.FamilyName, requestedStyle)
            .SelectMany(fallbackRun => {
                bool allowInstalledFace = webFonts.AllowInstalledFontFaces
                    && !webFonts.Slots.ContainsKey(fallbackRun.FamilyName);
                string family = ResolvePdfFontFamilyForText(
                    fallbackRun.FamilyName,
                    fallbackRun.Text,
                    visual.Font.IsBold,
                    visual.Font.IsItalic,
                    visual.FontDescriptor,
                    allowInstalledFace,
                    webFonts.Options);
                IReadOnlyList<NamedFaceStyleRun> faceRuns = allowInstalledFace
                    ? PlanNamedFaceStyleRuns(fallbackRun.Text, family,
                        visual.Font.IsBold, visual.Font.IsItalic, webFonts.Options)
                    : new[] { new NamedFaceStyleRun(fallbackRun.Text,
                        visual.Font.IsBold, visual.Font.IsItalic, true) };
                return faceRuns.Select(faceRun => new PdfCore.PdfTextRun(
                    faceRun.Text,
                    bold: faceRun.Bold,
                    underline: visual.Font.IsUnderline,
                    color: PdfCore.PdfColor.FromOfficeColorOrNull(visual.Color),
                    italic: faceRun.Italic,
                    strike: visual.Font.IsStrikethrough,
                    fontSize: visual.Font.Size * PointsPerCssPixel,
                    font: MapFont(family, faceRun.Text, requestedStyle, webFonts),
                    linkUri: link,
                    linkContents: link == null ? null : pdfText,
                    linkDestinationName: linkDestination,
                    fontFamily: family,
                    baseline: MapTextBaseline(visual.Baseline),
                    underlineStyle: visual.UnderlineStyle,
                    strikeStyle: visual.StrikethroughStyle,
                    decorationColor: PdfCore.PdfColor.FromOfficeColorOrNull(visual.DecorationColor))
                    .WithFeatureSettings(visual.FeatureSettings));
            }).ToArray();
        void Paint(PdfCore.PdfPageCanvas target) => target.PositionedText(
            runs,
            asSpan ? PdfCore.PdfCanvasTextStructureRole.Span : MapStructureRole(visual.SemanticRole),
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            frameWidth * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            PdfCore.PdfColor.FromOfficeColorOrNull(visual.Color),
            MapAlignment(visual.Alignment),
            baselineFontSize.Value * PointsPerCssPixel,
            visual.LineHeight * PointsPerCssPixel,
            (visual.TextPaintWidth ?? visual.TextAdvanceWidth) * PointsPerCssPixel,
            visual.PaintTopOverflow * PointsPerCssPixel,
            fontMetricScale: PointsPerCssPixel);
        double logicalWidth = visual.TextPaintWidth ?? visual.TextAdvanceWidth ?? frameWidth;
        double logicalHeight = Math.Max(0.01D, Math.Min(visual.Height, visual.Font.Size));
        var carrier = logicalClip?.ConstrainLogicalRectangle(visual.X, visual.Y, logicalWidth, logicalHeight)
            ?? (X: visual.X, Y: visual.Y, Width: logicalWidth, Height: logicalHeight);
        // Native text already carries transform-aware glyph geometry. Project
        // only page-space clipped native text; transformed logical/outlined
        // groups retain their own replacement carriers without changing the
        // native reader's existing partially-visible text policy.
        if (!logicalTextOwned && (!logicalClip.HasValue || logicalClip.Value.ConstrainToSurface)
            && (carrier.X != visual.X || carrier.Y != visual.Y
            || carrier.Width != logicalWidth || carrier.Height != logicalHeight)) {
            canvas.ActualText(pdfText, carrier.X * PointsPerCssPixel, carrier.Y * PointsPerCssPixel,
                carrier.Width * PointsPerCssPixel, carrier.Height * PointsPerCssPixel, Paint);
        } else Paint(canvas);
    }

}
