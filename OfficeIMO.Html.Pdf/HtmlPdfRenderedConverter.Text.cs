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
        bool constrainToSurface = true) {
        if (visual.Text.Length == 0) return;
        // Canvas and outline writers must anchor a script to the original line's metrics.
        baselineFontSize ??= visual.Font.Size;
        visual = visual.ResolveBaselineForPainting();
        if (visual.Y < 0D) {
            HtmlRenderText shifted = (HtmlRenderText)visual.TranslatePaint(0D, -visual.Y, visual.PaintOrder);
            canvas.Effect(OfficeTransform.Translate(0D, visual.Y * PointsPerCssPixel), 1D,
                nested => AddText(nested, shifted, webFonts, conversionReport, surfaceWidth, asSpan, logicalTextOwned, cancellationToken, baselineFontSize, suppressLink: suppressLink, preservePositionedFrame: preservePositionedFrame, constrainToSurface: constrainToSurface));
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
                    preservePositionedFrame)) {
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
                    asSpan, logicalTextOwned, cancellationToken, baselineFontSize, colorOpacityApplied: true, suppressLink: suppressLink, preservePositionedFrame: preservePositionedFrame, constrainToSurface: constrainToSurface));
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
        canvas.PositionedText(
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
    }

}
