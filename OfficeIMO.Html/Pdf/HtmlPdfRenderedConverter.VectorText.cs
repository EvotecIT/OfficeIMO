using OfficeIMO.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private const double SyntheticItalicShear = 0.22D;
    private static bool TryAddOutlinedText(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderText visual,
        RegisteredWebFonts webFonts,
        PdfCore.PdfConversionReport conversionReport,
        double frameWidth,
        bool asSpan,
        bool logicalTextOwned,
        CancellationToken cancellationToken,
        double baselineFontSize,
        bool suppressLink,
        bool preservePositionedFrame,
        ClipBounds? logicalClip) {
        cancellationToken.ThrowIfCancellationRequested();
        if (visual.PreparedShaping is OfficePreparedTextShaping prepared) {
            return TryAddPreparedOutlinedText(canvas, visual, prepared, webFonts, conversionReport,
                frameWidth, asSpan, logicalTextOwned, cancellationToken, baselineFontSize,
                suppressLink, preservePositionedFrame, logicalClip);
        }
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
                    run.FamilyRequest,
                    requestedStyle,
                    visual.Font.Size,
                    out OfficeFontFace? face)
                || face == null) {
                return false;
            }
            bool simulateBold = (requestedStyle & OfficeFontStyle.Bold) == OfficeFontStyle.Bold &&
                (face.Style & OfficeFontStyle.Bold) != OfficeFontStyle.Bold;
            bool simulateItalic = (requestedStyle & OfficeFontStyle.Italic) == OfficeFontStyle.Italic &&
                (face.Style & OfficeFontStyle.Italic) != OfficeFontStyle.Italic;
            resolvedRuns.Add(new OutlinedFontRun(run.Text, face, simulateBold, simulateItalic));
            // The managed PDF shaper can retain features in static embedded fonts. Keep
            // synthetic styles, configured providers and color fonts on their outline route;
            // substitutions can select color glyphs absent from the source character map.
            requiresOutlines |= (preservePositionedFrame && (simulateBold || simulateItalic)) ||
                !face.CanEmbedAsStaticPdfFont ||
                (!visual.FeatureSettings.IsDefault &&
                    (simulateBold || simulateItalic || webFonts.TextShapingProvider != null ||
                        (face.Program is OfficeTrueTypeFont colorFont && colorFont.HasColorGlyphs) ||
                        !webFonts.Slots.ContainsKey(run.FamilyName))) ||
                ContainsColorGlyph(face.Program, run.Text);
        }
        if (!requiresOutlines) {
            foreach (OutlinedFontRun run in resolvedRuns) {
                cancellationToken.ThrowIfCancellationRequested();
                if (RequiresProviderOwnedLayout(run.Face!, run.Text)) {
                    requiresOutlines = true;
                    break;
                }
            }
        }
        if (!requiresOutlines) return false;

        webFonts.OutlineBudget.ValidateTextLength(visual.Text.Length);
        if (webFonts.OutlineBudget.IsPathLimitReached)
            throw new InvalidOperationException("HTML-to-PDF outlined text exceeded the configured path-command budget.");
        if (resolvedRuns.Any(run => run.Program is not IOfficeBoundedFontProgram)) {
            throw new InvalidOperationException(
                "Provider-owned HTML-to-PDF text outlines require IOfficeBoundedFontProgram so cancellation and output limits remain enforceable.");
        }
        var runs = new List<OutlinedFontRun>(resolvedRuns.Count);
        foreach (OutlinedFontRun run in resolvedRuns) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeTextShapingResult? shapingResult = ShapeOutlinedRun(
                run,
                visual.FeatureSettings,
                webFonts,
                cancellationToken,
                out string? shapedText);
            if (shapingResult == null) {
                ReportDeclinedOutlinedShaping(webFonts, visual, run);
            }
            double advance = shapingResult == null
                ? run.Program.Measure(run.Text, visual.Font.Size)
                : run.Program.MeasureShapedText(shapedText!, shapingResult, visual.Font.Size);
            cancellationToken.ThrowIfCancellationRequested();
            if (double.IsNaN(advance) || double.IsInfinity(advance) || advance < 0D) return false;
            runs.Add(new OutlinedFontRun(
                run.Text,
                run.Face!,
                run.SimulateBold,
                run.SimulateItalic,
                advance,
                shapedText,
                shapingResult));
        }

        return TryPaintOutlinedText(canvas, visual, webFonts, conversionReport, frameWidth,
            asSpan, logicalTextOwned, cancellationToken, baselineFontSize, suppressLink,
            preservePositionedFrame, logicalClip, runs);
    }

    // Both freshly shaped and retained runs use the existing bounded contour,
    // frame, decoration, logical-text and link writer.
    private static bool TryPaintOutlinedText(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderText visual,
        RegisteredWebFonts webFonts,
        PdfCore.PdfConversionReport conversionReport,
        double frameWidth,
        bool asSpan,
        bool logicalTextOwned,
        CancellationToken cancellationToken,
        double baselineFontSize,
        bool suppressLink,
        bool preservePositionedFrame,
        ClipBounds? logicalClip,
        IReadOnlyList<OutlinedFontRun> runs) {
        double measuredAdvance = runs.Sum(run => run.Advance);
        if (measuredAdvance <= 0D) return false;
        double? paintWidth = visual.TextPaintWidth ?? visual.TextAdvanceWidth;
        double resolvedAdvance = paintWidth.HasValue && Math.Abs(paintWidth.Value) > 0.0001D
            ? Math.Abs(paintWidth.Value)
            : measuredAdvance;
        double scaleX = resolvedAdvance / measuredAdvance;
        double textX = ResolveOutlinedTextX(frameWidth, resolvedAdvance, visual.Alignment);
        var lineHeights = new double[runs.Count];
        var baselineOffsets = new double[runs.Count];
        for (int index = 0; index < runs.Count; index++) {
            IOfficeFontProgram program = runs[index].Program;
            lineHeights[index] = program.LineHeight(visual.Font.Size);
            baselineOffsets[index] = ResolveBaselineOffset(
                program,
                visual.Font.Size,
                lineHeights[index]);
        }
        // Script glyphs share the unscaled line's baseline before the authored displacement.
        OutlinedLineMetrics glyphMetrics = ResolveOutlinedLineMetrics(visual.Height, lineHeights, baselineOffsets);
        var lineMetrics = new OutlinedLineMetrics(
            glyphMetrics.TextTop + baselineFontSize - glyphMetrics.Baseline,
            baselineFontSize, glyphMetrics.LineHeight);
        var allContours = new List<List<OfficePoint>>();
        var paintGroups = new List<OutlinedPaintContours>();
        int retainedPointCount = 0;
        double cursor = textX;
        for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
            OutlinedFontRun run = runs[runIndex];
            cancellationToken.ThrowIfCancellationRequested();
            var bounded = (IOfficeBoundedFontProgram)run.Program;
            int availablePoints = webFonts.OutlineBudget.RemainingPointAllowance - retainedPointCount;
            if (availablePoints <= 0) {
                throw new InvalidOperationException("HTML-to-PDF outlined text exceeded the configured path-command budget.");
            }
            double runTop = lineMetrics.Baseline - baselineOffsets[runIndex];
            var sourcePaintGroups = new List<OutlinedPaintContours>();
            if (run.PreparedShaping is OfficePreparedTextShaping prepared) {
                sourcePaintGroups.Add(new OutlinedPaintContours(visual.Color,
                    prepared.GetContours(cursor, runTop, visual.Font.Size, availablePoints,
                        cancellationToken, webFonts.OutlineBudget.CffOperationBudget)));
            } else if (run.Program is OfficeTrueTypeFont trueType &&
                (run.ShapingResult == null
                    ? trueType.TryGetColorTextContours(
                        run.Text,
                        cursor,
                        runTop,
                        visual.Font.Size,
                        visual.FontPalette,
                        visual.Color,
                        availablePoints,
                        cancellationToken,
                        out List<OfficeColorGlyphContours> colorLayers)
                    : trueType.TryGetShapedColorTextContours(
                        run.ShapedText!,
                        run.ShapingResult,
                        cursor,
                        runTop,
                        visual.Font.Size,
                        visual.FontPalette,
                        visual.Color,
                        availablePoints,
                        cancellationToken,
                        out colorLayers))) {
                foreach (OfficeColorGlyphContours layer in colorLayers) {
                    sourcePaintGroups.Add(new OutlinedPaintContours(layer.Color, layer.Contours));
                }
            } else {
                List<List<OfficePoint>> contours;
                if (bounded is IOfficeCffBoundedFontProgram cff) {
                    contours = run.ShapingResult == null
                        ? cff.GetTextContoursBounded(
                            run.Text,
                            cursor,
                            runTop,
                            visual.Font.Size,
                            availablePoints,
                            cancellationToken,
                            webFonts.OutlineBudget.CffOperationBudget)
                        : cff.GetShapedTextContoursBounded(
                            run.ShapedText!,
                            run.ShapingResult,
                            cursor,
                            runTop,
                            visual.Font.Size,
                            availablePoints,
                            cancellationToken,
                            webFonts.OutlineBudget.CffOperationBudget);
                } else {
                    contours = run.ShapingResult == null
                        ? bounded.GetTextContoursBounded(
                            run.Text,
                            cursor,
                            runTop,
                            visual.Font.Size,
                            availablePoints,
                            cancellationToken)
                        : bounded.GetShapedTextContoursBounded(
                            run.ShapedText!,
                            run.ShapingResult,
                            cursor,
                            runTop,
                            visual.Font.Size,
                            availablePoints,
                            cancellationToken);
                }
                sourcePaintGroups.Add(new OutlinedPaintContours(visual.Color, contours));
            }
            cancellationToken.ThrowIfCancellationRequested();
            int runPointCount = 0;
            double italicBottom = runTop + (preservePositionedFrame ? visual.Font.Size : lineHeights[runIndex]);
            double boldOffset = preservePositionedFrame
                ? OfficeSyntheticTextStyle.BoldOffset(visual.Font.Size) : Math.Max(1D, visual.Font.Size / 22D);
            int copies = run.SimulateBold ? 2 : 1;
            foreach (OutlinedPaintContours sourceGroup in sourcePaintGroups) {
                var transformedGroup = new List<List<OfficePoint>>();
                foreach (List<OfficePoint> contour in sourceGroup.Contours) {
                    int retainedForContour = checked(contour.Count * copies);
                    if (retainedForContour > availablePoints - runPointCount) {
                        throw new InvalidOperationException("HTML-to-PDF outlined text exceeded the configured path-command budget.");
                    }
                    runPointCount += retainedForContour;
                    retainedPointCount += retainedForContour;
                    List<OfficePoint> transformed = contour.Select(point => TransformOutlinedPoint(
                        point,
                        textX,
                        scaleX,
                        italicBottom,
                        run.SimulateItalic,
                        0D, preservePositionedFrame)).ToList();
                    transformedGroup.Add(transformed);
                    allContours.Add(transformed);
                    if (run.SimulateBold) {
                        List<OfficePoint> bold = contour.Select(point => TransformOutlinedPoint(
                            point,
                            textX,
                            scaleX,
                            italicBottom,
                            run.SimulateItalic,
                            boldOffset, preservePositionedFrame)).ToList();
                        transformedGroup.Add(bold);
                        allContours.Add(bold);
                    }
                }
                if (transformedGroup.Count > 0) paintGroups.Add(new OutlinedPaintContours(sourceGroup.Color, transformedGroup));
            }
            cursor += run.Advance;
        }

        // A selected glyph such as a space can have advance and logical source
        // without ink. Keep its admitted result and ActualText; absence of contours
        // alone must not send it through a fresh text-shaping fallback.
        bool retainedEmptyRun = allContours.Count == 0 && runs.All(run => run.PreparedShaping != null);
        if (!TryResolveContourBounds(
                allContours,
                out double minimumX,
                out double minimumY,
                out double maximumX,
                out double maximumY)) {
            if (!retainedEmptyRun) return false;
            minimumX = textX;
            minimumY = lineMetrics.TextTop;
            maximumX = textX + resolvedAdvance;
            maximumY = minimumY + Math.Max(0.01D, visual.Font.Size);
        }
        OutlinedTextFrame frame = ResolveOutlinedTextFrame(
            visual.X,
            visual.Y,
            visual.Height,
            frameWidth,
            textX,
            minimumX,
            minimumY,
            maximumX,
            maximumY);
        double outlineOffsetX = frame.OffsetX;
        double outlineOffsetY = frame.OffsetY;
        double decorationThickness = Math.Max(0.5D, visual.Font.Size / 16D);
        var decorationCommands = new List<OfficePathCommand>();
        if (visual.Font.IsUnderline) {
            AppendRectangle(
                decorationCommands,
                textX + outlineOffsetX,
                lineMetrics.TextTop + (lineMetrics.LineHeight * 0.86D) + outlineOffsetY,
                resolvedAdvance,
                decorationThickness,
                webFonts.OutlineBudget);
        }
        if (visual.Font.IsStrikethrough) {
            AppendRectangle(
                decorationCommands,
                textX + outlineOffsetX,
                lineMetrics.TextTop + (lineMetrics.LineHeight * 0.52D) + outlineOffsetY,
                resolvedAdvance,
                decorationThickness,
                webFonts.OutlineBudget);
        }
        if (paintGroups.Count == 0 && decorationCommands.Count == 0 && !retainedEmptyRun) return false;

        double drawingWidth = frame.Width;
        double drawingHeight = frame.Height;
        var drawing = new OfficeDrawing(drawingWidth, drawingHeight);
        foreach (OutlinedPaintContours group in paintGroups) {
            var commands = new List<OfficePathCommand>();
            AppendContours(commands, group.Contours, outlineOffsetX, outlineOffsetY, webFonts.OutlineBudget, cancellationToken);
            if (commands.Count == 0) continue;
            var path = OfficeShape.Path(drawingWidth, drawingHeight, commands);
            path.FillColor = group.Color;
            path.FillRule = OfficeFillRule.NonZero;
            path.StrokeColor = null;
            drawing.AddShape(path, 0D, 0D);
        }
        if (decorationCommands.Count > 0) {
            var decoration = OfficeShape.Path(drawingWidth, drawingHeight, decorationCommands);
            decoration.FillColor = visual.DecorationColor;
            decoration.FillRule = OfficeFillRule.NonZero;
            decoration.StrokeColor = null;
            drawing.AddShape(decoration, 0D, 0D);
        }
        string? link = suppressLink || string.IsNullOrWhiteSpace(visual.Text) || IsFragmentLink(visual.LinkUri) ? null : visual.LinkUri;
        string? linkDestination = !suppressLink && IsFragmentLink(visual.LinkUri)
            ? MapNamedDestination(visual.LinkUri!.Substring(1))
            : null;
        Action<PdfCore.PdfPageCanvas> addDrawing = target => {
            if (drawing.Elements.Count == 0) {
                // Logical groups still need a content item when the selected glyphs
                // have no ink. The transparent frame also retains per-visual links.
                OfficeShape area = OfficeShape.Rectangle(drawingWidth * PointsPerCssPixel, drawingHeight * PointsPerCssPixel);
                area.FillColor = null;
                area.StrokeColor = null;
                target.Shape(area, frame.X * PointsPerCssPixel, frame.Y * PointsPerCssPixel,
                    style: new PdfCore.PdfDrawingStyle { Decorative = true }, linkUri: link, linkContents: link == null ? null : visual.Text);
                return;
            }
            target.Drawing(
                drawing,
                frame.X * PointsPerCssPixel,
                frame.Y * PointsPerCssPixel,
                drawingWidth * PointsPerCssPixel,
                drawingHeight * PointsPerCssPixel,
                style: new PdfCore.PdfDrawingStyle { Decorative = true },
                linkUri: link,
                linkContents: link == null ? null : visual.Text);
        };
        Action<PdfCore.PdfPageCanvas> addDrawingAndLink = target => {
            addDrawing(target);
            if (linkDestination != null) {
                target.LinkToNamedDestination(
                    linkDestination,
                    frame.X * PointsPerCssPixel,
                    frame.Y * PointsPerCssPixel,
                    drawingWidth * PointsPerCssPixel,
                    drawingHeight * PointsPerCssPixel,
                    visual.Text);
            }
        };

        if (logicalTextOwned) {
            addDrawingAndLink(canvas);
        } else {
            double logicalHeight = Math.Max(0.01D, Math.Min(visual.Height, visual.Font.Size));
            var carrier = logicalClip?.ConstrainLogicalRectangle(visual.X + textX, visual.Y, resolvedAdvance, logicalHeight)
                ?? (X: visual.X + textX, Y: visual.Y, Width: resolvedAdvance, Height: logicalHeight);
            PdfCore.PdfCanvasTextStructureRole role = asSpan
                ? PdfCore.PdfCanvasTextStructureRole.Span
                : MapStructureRole(visual.SemanticRole);
            if (role == PdfCore.PdfCanvasTextStructureRole.Span) {
                canvas.ActualText(
                    visual.Text,
                    carrier.X * PointsPerCssPixel,
                    carrier.Y * PointsPerCssPixel,
                    carrier.Width * PointsPerCssPixel,
                    carrier.Height * PointsPerCssPixel,
                    addDrawingAndLink, logicalOrder: visual.SemanticSourceOrder);
            } else {
                canvas.Structure(MapOutlinedTextStructureRole(role), nested =>
                    nested.ActualText(
                        visual.Text,
                        carrier.X * PointsPerCssPixel,
                        carrier.Y * PointsPerCssPixel,
                        carrier.Width * PointsPerCssPixel,
                        carrier.Height * PointsPerCssPixel,
                        addDrawingAndLink, logicalOrder: visual.SemanticSourceOrder),
                    new PdfCore.PdfCanvasStructureOptions { LogicalOrder = visual.SemanticSourceOrder });
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

    private static bool ContainsColorGlyph(IOfficeFontProgram program, string text) {
        if (program is not IOfficeColorFontProgram colorProgram) return false;
        for (int index = 0; index < text.Length;) {
            int scalar;
            char first = text[index++];
            if (char.IsHighSurrogate(first) && index < text.Length && char.IsLowSurrogate(text[index])) {
                scalar = char.ConvertToUtf32(first, text[index++]);
            } else {
                scalar = first;
            }
            if (program.TryGetGlyphMetrics(scalar, out int glyphId, out _) && colorProgram.HasColorGlyph(glyphId)) return true;
        }
        return false;
    }

    private static OfficeTextShapingResult? ShapeOutlinedRun(
        OutlinedFontRun run,
        OfficeTextFeatureSettings featureSettings,
        RegisteredWebFonts webFonts,
        CancellationToken cancellationToken,
        out string? shapedText) {
        shapedText = null;
        IOfficeTextShapingProvider? provider = webFonts.TextShapingProvider;
        if (provider == null && !featureSettings.IsDefault) provider = OfficeManagedTextShapingProvider.Instance;
        if (provider == null || run.Program.ProvidesComplexTextLayout) return null;
        shapedText = OfficeArabicTextShaper.ToLogicalText(run.Text);
        cancellationToken.ThrowIfCancellationRequested();
        IOfficeFontProgram program = run.Program;
        OfficeTextShapingResult? shapingResult = provider.ShapeText(new OfficeTextShapingRequest(
            shapedText,
            program.DisplayName ?? run.Face!.FamilyName,
            program.GetFontDataForShaping(),
            program.IsOpenTypeCff,
            program.UnitsPerEm,
            OfficeTextElements.ResolveBaseDirection(shapedText),
            webFonts.TextShapingLanguage,
            cancellationToken,
            program.CollectionIndex,
            run.Face!.VariationCoordinatesForShaping,
            cloneFontData: false,
            fontProgramCacheKey: program,
            featureSettings: featureSettings));
        if (shapingResult == null) shapedText = null;
        return shapingResult;
    }

    private static void ReportDeclinedOutlinedShaping(
        RegisteredWebFonts webFonts,
        HtmlRenderText visual,
        OutlinedFontRun run) {
        if (webFonts.TextShapingProvider == null
            || run.Program.ProvidesComplexTextLayout
            || !RequiresConfiguredTextShaping(run.Text)) return;
        string source = visual.Source ?? "html-text";
        if (webFonts.Diagnostics.Any(diagnostic =>
                string.Equals(diagnostic.Code, HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported, StringComparison.Ordinal)
                && string.Equals(diagnostic.Source, source, StringComparison.Ordinal))) return;
        webFonts.Diagnostics.Add(
            "OfficeIMO.Html.Pdf",
            HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported,
            "A complex-script outlined run required provider-owned shaping, but the configured provider declined it; scalar outlines were used.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "provider-declined",
            OfficeConversionLossKind.Approximation);
    }

    private static bool RequiresConfiguredTextShaping(string text) =>
        OfficeTextElements.ContainsShapingRequiredScript(text)
        || OfficeTextElements.ContainsVariationSelector(text)
        || OfficeTextElements.ContainsCombiningMarkOrJoiner(text)
        || OfficeTextElements.ContainsJoiningScript(text)
            && !OfficeArabicTextShaper.CanShapeAllJoiningCharacters(text);

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

    private static double ResolveBaselineOffset(
        IOfficeFontProgram program,
        double fontSize,
        double lineHeight) {
        double offset = program is IOfficeFontBaselineMetrics metrics
            ? metrics.BaselineOffset(fontSize)
            : lineHeight * 0.8D;
        if (double.IsNaN(offset) || double.IsInfinity(offset)) return lineHeight * 0.8D;
        return Math.Max(0D, Math.Min(lineHeight, offset));
    }

    internal static OutlinedLineMetrics ResolveOutlinedLineMetrics(
        double visualHeight,
        IReadOnlyList<double> lineHeights,
        IReadOnlyList<double> baselineOffsets) {
        if (lineHeights == null) throw new ArgumentNullException(nameof(lineHeights));
        if (baselineOffsets == null) throw new ArgumentNullException(nameof(baselineOffsets));
        if (lineHeights.Count == 0 || lineHeights.Count != baselineOffsets.Count) {
            throw new ArgumentException("Outlined line metrics require one baseline offset per font run.");
        }

        double maximumAscent = 0D;
        double maximumDescent = 0D;
        for (int index = 0; index < lineHeights.Count; index++) {
            double lineHeight = lineHeights[index];
            double baselineOffset = baselineOffsets[index];
            if (double.IsNaN(lineHeight) || double.IsInfinity(lineHeight) || lineHeight <= 0D ||
                double.IsNaN(baselineOffset) || double.IsInfinity(baselineOffset) ||
                baselineOffset < 0D || baselineOffset > lineHeight) {
                throw new ArgumentOutOfRangeException(nameof(lineHeights), "Outlined font metrics must be finite and bounded by the line height.");
            }
            maximumAscent = Math.Max(maximumAscent, baselineOffset);
            maximumDescent = Math.Max(maximumDescent, lineHeight - baselineOffset);
        }

        double combinedLineHeight = maximumAscent + maximumDescent;
        double textTop = Math.Max(0D, (visualHeight - combinedLineHeight) / 2D);
        return new OutlinedLineMetrics(textTop, textTop + maximumAscent, combinedLineHeight);
    }

    internal static OutlinedTextFrame ResolveOutlinedTextFrame(
        double visualX,
        double visualY,
        double visualHeight,
        double frameWidth,
        double textX,
        double minimumX,
        double minimumY,
        double maximumX,
        double maximumY) {
        double offsetX = minimumX < textX ? textX - minimumX : 0D;
        double offsetY = minimumY < 0D ? -minimumY : 0D;
        return new OutlinedTextFrame(
            visualX - offsetX,
            visualY - offsetY,
            Math.Max(0.01D, Math.Max(frameWidth + offsetX, maximumX + offsetX)),
            Math.Max(0.01D, Math.Max(visualHeight + offsetY, maximumY + offsetY)),
            offsetX,
            offsetY);
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

    private static OfficePoint TransformOutlinedPoint(
        OfficePoint point,
        double originX,
        double scaleX,
        double bottom,
        bool simulateItalic,
        double boldOffset, bool preservePositionedFrame) {
        if (preservePositionedFrame) {
            double positionedX = originX + ((point.X - originX) * scaleX) + boldOffset;
            if (simulateItalic) positionedX += OfficeSyntheticTextStyle.ItalicOffset(bottom, point.Y);
            return new OfficePoint(positionedX, point.Y);
        }
        double x = point.X + boldOffset;
        if (simulateItalic) x += (bottom - point.Y) * SyntheticItalicShear;
        return new OfficePoint(originX + ((x - originX) * scaleX), point.Y);
    }

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
        foreach (OfficeFontFace face in runs.Select(run => run.Face).OfType<OfficeFontFace>().Distinct()) {
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

    private sealed class OutlinedPaintContours {
        internal OutlinedPaintContours(OfficeColor color, List<List<OfficePoint>> contours) {
            Color = color;
            Contours = contours;
        }

        internal OfficeColor Color { get; }
        internal List<List<OfficePoint>> Contours { get; }
    }

    internal readonly struct OutlinedTextFrame {
        internal OutlinedTextFrame(
            double x,
            double y,
            double width,
            double height,
            double offsetX,
            double offsetY) {
            X = x;
            Y = y;
            Width = width;
            Height = height;
            OffsetX = offsetX;
            OffsetY = offsetY;
        }

        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal double OffsetX { get; }
        internal double OffsetY { get; }
    }

    internal readonly struct OutlinedLineMetrics {
        internal OutlinedLineMetrics(double textTop, double baseline, double lineHeight) {
            TextTop = textTop;
            Baseline = baseline;
            LineHeight = lineHeight;
        }

        internal double TextTop { get; }
        internal double Baseline { get; }
        internal double LineHeight { get; }
    }
}
