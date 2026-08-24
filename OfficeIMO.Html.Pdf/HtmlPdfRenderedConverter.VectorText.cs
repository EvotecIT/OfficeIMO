using OfficeIMO.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static bool TryAddOutlinedText(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderText visual,
        RegisteredWebFonts webFonts,
        PdfCore.PdfConversionReport conversionReport,
        double frameWidth,
        bool asSpan,
        bool logicalTextOwned,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeFontStyle requestedStyle = (visual.Font.IsBold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (visual.Font.IsItalic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        IReadOnlyList<OfficeFontFallbackRun> planned = webFonts.Faces.PlanFallbackRuns(
            visual.Text,
            visual.Font.FamilyName,
            requestedStyle);
        if (planned.Count == 0) return false;

        var resolvedRuns = new List<OutlinedFontRun>(planned.Count);
        bool requiresOutlines = false;
        foreach (OfficeFontFallbackRun run in planned) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!webFonts.Faces.TryResolveFaceForText(
                    run.Text,
                    run.FamilyName,
                    requestedStyle,
                    out OfficeFontFace? face)
                || face == null) {
                return false;
            }
            resolvedRuns.Add(new OutlinedFontRun(run.Text, face));
            requiresOutlines |= !face.CanEmbedAsStaticPdfFont;
        }
        if (!requiresOutlines) {
            foreach (OutlinedFontRun run in resolvedRuns) {
                cancellationToken.ThrowIfCancellationRequested();
                if (RequiresProviderOwnedLayout(run.Face, run.Text)) {
                    requiresOutlines = true;
                    break;
                }
            }
        }
        if (!requiresOutlines) return false;

        webFonts.OutlineBudget.ValidateTextLength(visual.Text.Length);
        if (resolvedRuns.Any(run => run.Face.Program is not IOfficeBoundedFontProgram)) {
            throw new InvalidOperationException(
                "Provider-owned HTML-to-PDF text outlines require IOfficeBoundedFontProgram so cancellation and output limits remain enforceable.");
        }
        var runs = new List<OutlinedFontRun>(resolvedRuns.Count);
        foreach (OutlinedFontRun run in resolvedRuns) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeTextShapingResult? shapingResult = ShapeOutlinedRun(
                run,
                webFonts,
                cancellationToken,
                out string? shapedText);
            double advance = shapingResult == null
                ? run.Face.Program.Measure(run.Text, visual.Font.Size)
                : run.Face.Program.MeasureShapedText(shapedText!, shapingResult, visual.Font.Size);
            cancellationToken.ThrowIfCancellationRequested();
            if (double.IsNaN(advance) || double.IsInfinity(advance) || advance < 0D) return false;
            runs.Add(new OutlinedFontRun(run.Text, run.Face, advance, shapedText, shapingResult));
        }

        double measuredAdvance = runs.Sum(run => run.Advance);
        if (measuredAdvance <= 0D) return false;
        double resolvedAdvance = visual.TextAdvanceWidth.HasValue
            && Math.Abs(visual.TextAdvanceWidth.Value) > 0.0001D
                ? Math.Abs(visual.TextAdvanceWidth.Value)
                : measuredAdvance;
        double scaleX = resolvedAdvance / measuredAdvance;
        double textX = ResolveOutlinedTextX(frameWidth, resolvedAdvance, visual.Alignment);
        double lineHeight = runs.Max(run => run.Face.Program.LineHeight(visual.Font.Size));
        double textTop = Math.Max(0D, (visual.Height - lineHeight) / 2D);
        var allContours = new List<List<OfficePoint>>();
        int retainedPointCount = 0;
        double cursor = textX;
        foreach (OutlinedFontRun run in runs) {
            cancellationToken.ThrowIfCancellationRequested();
            var bounded = (IOfficeBoundedFontProgram)run.Face.Program;
            int availablePoints = webFonts.OutlineBudget.RemainingPointAllowance - retainedPointCount;
            if (availablePoints <= 0) {
                throw new InvalidOperationException("HTML-to-PDF outlined text exceeded the configured path-command budget.");
            }
            List<List<OfficePoint>> contours = run.ShapingResult == null
                ? bounded.GetTextContoursBounded(
                    run.Text,
                    cursor,
                    textTop,
                    visual.Font.Size,
                    availablePoints,
                    cancellationToken)
                : bounded.GetShapedTextContoursBounded(
                    run.ShapedText!,
                    run.ShapingResult,
                    cursor,
                    textTop,
                    visual.Font.Size,
                    availablePoints,
                    cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            int runPointCount = 0;
            foreach (List<OfficePoint> contour in contours) {
                if (contour.Count > availablePoints - runPointCount) {
                    throw new InvalidOperationException("HTML-to-PDF outlined text exceeded the configured path-command budget.");
                }
                runPointCount += contour.Count;
                retainedPointCount += contour.Count;
                allContours.Add(contour.Select(point => ScalePointX(point, textX, scaleX)).ToList());
            }
            cursor += run.Advance;
        }

        if (!TryResolveContourBounds(
                allContours,
                out double minimumX,
                out double minimumY,
                out double maximumX,
                out double maximumY)) {
            return false;
        }
        double outlineOffsetX = minimumX < textX ? textX - minimumX : 0D;
        double outlineOffsetY = minimumY < 0D ? -minimumY : 0D;
        var commands = new List<OfficePathCommand>();
        AppendContours(
            commands,
            allContours,
            outlineOffsetX,
            outlineOffsetY,
            webFonts.OutlineBudget,
            cancellationToken);

        double decorationThickness = Math.Max(0.5D, visual.Font.Size / 16D);
        if (visual.Font.IsUnderline) {
            AppendRectangle(
                commands,
                textX,
                textTop + (lineHeight * 0.86D),
                resolvedAdvance,
                decorationThickness,
                webFonts.OutlineBudget);
        }
        if (visual.Font.IsStrikethrough) {
            AppendRectangle(
                commands,
                textX,
                textTop + (lineHeight * 0.52D),
                resolvedAdvance,
                decorationThickness,
                webFonts.OutlineBudget);
        }
        if (commands.Count == 0) return false;

        double drawingWidth = Math.Max(0.01D, Math.Max(frameWidth, maximumX + outlineOffsetX));
        double drawingHeight = Math.Max(0.01D, Math.Max(visual.Height, maximumY + outlineOffsetY));
        var path = OfficeShape.Path(drawingWidth, drawingHeight, commands);
        path.FillColor = visual.Color;
        path.FillRule = OfficeFillRule.EvenOdd;
        path.StrokeColor = null;
        var drawing = new OfficeDrawing(drawingWidth, drawingHeight)
            .AddShape(path, 0D, 0D);
        string? link = string.IsNullOrWhiteSpace(visual.Text) ? null : visual.LinkUri;
        Action<PdfCore.PdfPageCanvas> addDrawing = target => target.Drawing(
            drawing,
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            drawingWidth * PointsPerCssPixel,
            drawingHeight * PointsPerCssPixel,
            style: new PdfCore.PdfDrawingStyle { Decorative = true },
            linkUri: link,
            linkContents: link == null ? null : visual.Text);

        if (logicalTextOwned) {
            addDrawing(canvas);
        } else {
            PdfCore.PdfCanvasTextStructureRole role = asSpan
                ? PdfCore.PdfCanvasTextStructureRole.Span
                : MapStructureRole(visual.SemanticRole);
            if (role == PdfCore.PdfCanvasTextStructureRole.Span) {
                canvas.ActualText(
                    visual.Text,
                    visual.X * PointsPerCssPixel,
                    (visual.Y + Math.Min(visual.Height, visual.Font.Size)) * PointsPerCssPixel,
                    addDrawing);
            } else {
                canvas.Structure(MapOutlinedTextStructureRole(role), nested =>
                    nested.ActualText(
                        visual.Text,
                        visual.X * PointsPerCssPixel,
                        (visual.Y + Math.Min(visual.Height, visual.Font.Size)) * PointsPerCssPixel,
                        addDrawing));
            }
        }

        ReportOutlinedFontEvidence(conversionReport, visual, runs);
        return true;
    }

    private static bool RequiresProviderOwnedLayout(OfficeFontFace face, string text) {
        if (!face.CanEmbedAsStaticPdfFont || !face.Program.ProvidesComplexTextLayout) return false;
        try {
            return PdfCore.PdfTextDiagnostics.AnalyzeAdvancedTextLayout(
                text,
                face.Data,
                source: "OfficeIMO.Html.Pdf",
                fontName: face.Program.DisplayName).Count > 0;
        } catch (Exception exception) when (!(exception is OutOfMemoryException)) {
            // A face marked static-embeddable has already been parsed by the PDF font path. If a
            // later diagnostic probe cannot inspect it, keep the normal embedded-font behavior;
            // the PDF conversion report will retain the writer's exact font diagnostic.
            return false;
        }
    }

    private static OfficeTextShapingResult? ShapeOutlinedRun(
        OutlinedFontRun run,
        RegisteredWebFonts webFonts,
        CancellationToken cancellationToken,
        out string? shapedText) {
        shapedText = null;
        if (webFonts.TextShapingProvider == null || run.Face.Program.ProvidesComplexTextLayout) return null;
        shapedText = OfficeArabicTextShaper.ToLogicalText(run.Text);
        cancellationToken.ThrowIfCancellationRequested();
        IOfficeFontProgram program = run.Face.Program;
        OfficeTextShapingResult? shapingResult = webFonts.TextShapingProvider.ShapeText(new OfficeTextShapingRequest(
            shapedText,
            program.DisplayName ?? run.Face.FamilyName,
            program.GetFontDataForShaping(),
            program.IsOpenTypeCff,
            program.UnitsPerEm,
            OfficeTextElements.ResolveBaseDirection(shapedText),
            webFonts.TextShapingLanguage,
            cancellationToken,
            program.CollectionIndex,
            run.Face.VariationCoordinatesForShaping));
        if (shapingResult == null) shapedText = null;
        return shapingResult;
    }

    private static double ResolveOutlinedTextX(
        double frameWidth,
        double resolvedAdvance,
        OfficeTextAlignment alignment) {
        if (alignment == OfficeTextAlignment.Center) {
            return Math.Max(0D, (frameWidth - resolvedAdvance) / 2D);
        }
        if (alignment == OfficeTextAlignment.Right) {
            return Math.Max(0D, frameWidth - resolvedAdvance);
        }
        return 0D;
    }

    private static void AppendContours(
        ICollection<OfficePathCommand> commands,
        IEnumerable<List<OfficePoint>> contours,
        double offsetX,
        double offsetY,
        OutlinedTextBudget budget,
        CancellationToken cancellationToken) {
        foreach (List<OfficePoint> contour in contours) {
            cancellationToken.ThrowIfCancellationRequested();
            if (contour.Count < 3) continue;
            OfficePoint first = TranslatePoint(contour[0], offsetX, offsetY);
            budget.ConsumePathCommand();
            commands.Add(OfficePathCommand.MoveTo(first));
            int lastIndex = contour.Count - 1;
            if (contour[lastIndex] == contour[0]) lastIndex--;
            for (int index = 1; index <= lastIndex; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                budget.ConsumePathCommand();
                commands.Add(OfficePathCommand.LineTo(TranslatePoint(contour[index], offsetX, offsetY)));
            }
            budget.ConsumePathCommand();
            commands.Add(OfficePathCommand.Close());
        }
    }

    private static OfficePoint ScalePointX(OfficePoint point, double originX, double scaleX) =>
        new(originX + ((point.X - originX) * scaleX), point.Y);

    private static OfficePoint TranslatePoint(OfficePoint point, double offsetX, double offsetY) =>
        new(point.X + offsetX, point.Y + offsetY);

    private static bool TryResolveContourBounds(
        IEnumerable<List<OfficePoint>> contours,
        out double minimumX,
        out double minimumY,
        out double maximumX,
        out double maximumY) {
        minimumX = double.PositiveInfinity;
        minimumY = double.PositiveInfinity;
        maximumX = double.NegativeInfinity;
        maximumY = double.NegativeInfinity;
        bool found = false;
        foreach (List<OfficePoint> contour in contours) {
            foreach (OfficePoint point in contour) {
                found = true;
                minimumX = Math.Min(minimumX, point.X);
                minimumY = Math.Min(minimumY, point.Y);
                maximumX = Math.Max(maximumX, point.X);
                maximumY = Math.Max(maximumY, point.Y);
            }
        }
        return found;
    }

    private static void AppendRectangle(
        ICollection<OfficePathCommand> commands,
        double x,
        double y,
        double width,
        double height,
        OutlinedTextBudget budget) {
        if (width <= 0D || height <= 0D) return;
        budget.ConsumePathCommand();
        commands.Add(OfficePathCommand.MoveTo(x, y));
        budget.ConsumePathCommand();
        commands.Add(OfficePathCommand.LineTo(x + width, y));
        budget.ConsumePathCommand();
        commands.Add(OfficePathCommand.LineTo(x + width, y + height));
        budget.ConsumePathCommand();
        commands.Add(OfficePathCommand.LineTo(x, y + height));
        budget.ConsumePathCommand();
        commands.Add(OfficePathCommand.Close());
    }

    private static PdfCore.PdfCanvasStructureRole MapOutlinedTextStructureRole(
        PdfCore.PdfCanvasTextStructureRole role) {
        if (role == PdfCore.PdfCanvasTextStructureRole.Heading1) return PdfCore.PdfCanvasStructureRole.Heading1;
        if (role == PdfCore.PdfCanvasTextStructureRole.Heading2) return PdfCore.PdfCanvasStructureRole.Heading2;
        if (role == PdfCore.PdfCanvasTextStructureRole.Heading3) return PdfCore.PdfCanvasStructureRole.Heading3;
        if (role == PdfCore.PdfCanvasTextStructureRole.Heading4) return PdfCore.PdfCanvasStructureRole.Heading4;
        if (role == PdfCore.PdfCanvasTextStructureRole.Heading5) return PdfCore.PdfCanvasStructureRole.Heading5;
        if (role == PdfCore.PdfCanvasTextStructureRole.Heading6) return PdfCore.PdfCanvasStructureRole.Heading6;
        return PdfCore.PdfCanvasStructureRole.Paragraph;
    }

    private static void ReportOutlinedFontEvidence(
        PdfCore.PdfConversionReport conversionReport,
        HtmlRenderText visual,
        IEnumerable<OutlinedFontRun> runs) {
        foreach (OfficeFontFace face in runs.Select(run => run.Face).Distinct()) {
            bool alreadyReported = conversionReport.Warnings.Any(warning =>
                string.Equals(warning.Code, HtmlPdfDiagnosticCodes.FontProgramOutlined, StringComparison.Ordinal)
                && warning.Details.TryGetValue("ResourceFamily", out string? family)
                && string.Equals(family, face.ResourceFamilyName, StringComparison.Ordinal));
            if (alreadyReported) continue;
            conversionReport.Add(new PdfCore.PdfConversionWarning(
                "OfficeIMO.Html.Pdf",
                HtmlPdfDiagnosticCodes.FontProgramOutlined,
                visual.Source ?? "html-text",
                "A shaped font run was painted as vector outlines with logical ActualText retained for extraction and accessibility.",
                PdfCore.PdfConversionWarningSeverity.Information,
                details: new Dictionary<string, string> {
                    ["Family"] = face.FamilyName,
                    ["ResourceFamily"] = face.ResourceFamilyName,
                    ["Container"] = face.ContainerFormat.ToString(),
                    ["Program"] = face.Program.DisplayName ?? string.Empty,
                    ["Representation"] = "vector-outlines-plus-actual-text",
                    ["StaticPdfEmbeddable"] = face.CanEmbedAsStaticPdfFont ? "true" : "false"
                }));
        }
    }

    private readonly struct OutlinedFontRun {
        internal OutlinedFontRun(string text, OfficeFontFace face)
            : this(text, face, 0D, null, null) {
        }

        internal OutlinedFontRun(
            string text,
            OfficeFontFace face,
            double advance,
            string? shapedText,
            OfficeTextShapingResult? shapingResult) {
            Text = text;
            Face = face;
            Advance = advance;
            ShapedText = shapedText;
            ShapingResult = shapingResult;
        }

        internal string Text { get; }
        internal OfficeFontFace Face { get; }
        internal double Advance { get; }
        internal string? ShapedText { get; }
        internal OfficeTextShapingResult? ShapingResult { get; }
    }
}
