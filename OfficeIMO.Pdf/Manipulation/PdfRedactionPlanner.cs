using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Builds redaction impact previews without modifying the PDF.</summary>
internal static partial class PdfRedactionPlanner {
    private const double DefaultTextHeight = 12D;

    /// <summary>Plans redaction impact for bounded rectangle or exact path geometry.</summary>
    public static PdfRedactionPlan Plan(byte[] pdf, IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null, CancellationToken cancellationToken = default) {
        return Plan(pdf, areas, layoutOptions, options, includeHiddenOptionalContentText: false, excludeGeneratedRedactionMarks: false, cancellationToken);
    }

    private static PdfRedactionPlan CreateUnsupportedTextMappingPlan(byte[] pdf, PdfUnsupportedTextMappingException error,
        PdfLoadOptions? options, string code, CancellationToken cancellationToken) =>
        new(PdfInspector.Preflight(pdf, options, cancellationToken), Array.Empty<PdfRedactionArea>(),
            Array.Empty<PdfRedactionMatch>(), new[] { new PdfDiagnosticFinding(PdfDiagnosticSeverity.Error, code, error.Message) },
            Array.Empty<string>(), PdfRedactionPlan.ComputeSourceSha256(pdf));

    internal static PdfRedactionPlan PlanForVerification(byte[] pdf, IEnumerable<PdfRedactionArea> areas, PdfLoadOptions? options, CancellationToken cancellationToken = default) {
        return Plan(pdf, areas, layoutOptions: null, options, includeHiddenOptionalContentText: true, excludeGeneratedRedactionMarks: true, cancellationToken);
    }

    private static PdfRedactionPlan Plan(
        byte[] pdf,
        IEnumerable<PdfRedactionArea> areas,
        PdfTextLayoutOptions? layoutOptions,
        PdfLoadOptions? options,
        bool includeHiddenOptionalContentText,
        bool excludeGeneratedRedactionMarks,
        CancellationToken cancellationToken) {
        try {
            return PlanCore(pdf, areas, layoutOptions, options, includeHiddenOptionalContentText,
                excludeGeneratedRedactionMarks, cancellationToken);
        } catch (PdfUnsupportedTextMappingException error) {
            return CreateUnsupportedTextMappingPlan(pdf, error, options,
                "redaction-source-text-mapping-unsupported", cancellationToken);
        }
    }

    private static PdfRedactionPlan PlanCore(
        byte[] pdf,
        IEnumerable<PdfRedactionArea> areas,
        PdfTextLayoutOptions? layoutOptions,
        PdfLoadOptions? options,
        bool includeHiddenOptionalContentText,
        bool excludeGeneratedRedactionMarks,
        CancellationToken cancellationToken) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(areas, nameof(areas));
        cancellationToken.ThrowIfCancellationRequested();

        PdfRedactionArea[] areaArray = areas.ToArray();
        cancellationToken.ThrowIfCancellationRequested();
        if (areaArray.Length == 0) {
            throw new ArgumentException("At least one redaction area is required.", nameof(areas));
        }

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        var findings = new List<PdfDiagnosticFinding>();
        if (!preflight.CanReadLogicalObjects) {
            foreach (string message in preflight.GetCapabilityDiagnostics(PdfPreflightCapability.ReadLogicalObjects)) {
                findings.Add(new PdfDiagnosticFinding(PdfDiagnosticSeverity.Error, "redaction-plan-blocked", message));
            }

            return new PdfRedactionPlan(
                preflight,
                areaArray,
                Array.Empty<PdfRedactionMatch>(),
                findings.AsReadOnly(),
                searchCriteria: null,
                PdfRedactionPlan.ComputeSourceSha256(pdf));
        }

        PdfReadDocument readDocument = PdfReadDocument.Open(pdf, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<IReadOnlyList<PdfRedactionTextObjectScope>> reviewedTextObjectScopes = PdfRedactionPlan.CaptureReviewedTextObjectScopes(readDocument, areaArray);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<string> pageIdentities = PdfRedactionPlan.CapturePageIdentities(readDocument, areaArray, reviewedTextObjectScopes);
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocumentReadResult logical = PdfDocumentReadResult.From(readDocument, layoutOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocumentInfo info = preflight.UncheckedDocumentInfo ?? PdfInspector.Inspect(pdf, options);
        var matches = new List<PdfRedactionMatch>();
        var nestedPathPrimitivesByPage = new Dictionary<int, IReadOnlyList<PdfPageVisualPrimitive>>();
        var hiddenTextSpansByPage = new Dictionary<int, IReadOnlyList<PdfTextSpan>>();
        var glyphTextSpansByPage = new Dictionary<int, IReadOnlyList<PdfTextSpan>>();
        var hiddenImagePlacementsByPage = new Dictionary<int, IReadOnlyList<PdfImagePlacement>>();
        var inconclusiveOptionalContentPages = new HashSet<int>();
        var staleSelectionPages = new HashSet<int>();
        var generatedMarksByPage = new Dictionary<int, HashSet<int>>();

        PdfReadDocument? glyphDocument = null;
        foreach (PdfRedactionArea area in areaArray) {
            cancellationToken.ThrowIfCancellationRequested();
            if (area.GlyphSelection is PdfRedactionGlyphSelection selection && area.PageNumber <= readDocument.Pages.Count) {
                if (!glyphTextSpansByPage.TryGetValue(area.PageNumber, out IReadOnlyList<PdfTextSpan>? glyphSpans)) {
                    // A selection can name glyphs in hidden layers and artifacts, so resolve it against all of them.
                    glyphDocument ??= PdfReadDocument.Open(pdf, PdfLoadOptions.WithArtifactText(options), cancellationToken);
                    glyphSpans = glyphDocument.Pages[area.PageNumber - 1].GetGlyphTextSpans(
                        includeHiddenOptionalContent: true, cancellationToken);
                    glyphTextSpansByPage.Add(area.PageNumber, glyphSpans);
                }
                bool resolved = AddSelectedGlyphMatches(area, selection, glyphSpans, matches, cancellationToken);
                // Verification plans run on the rewritten output, where selected glyphs are expected to be gone.
                if (!resolved && !excludeGeneratedRedactionMarks && staleSelectionPages.Add(area.PageNumber)) {
                    findings.Add(new PdfDiagnosticFinding(PdfDiagnosticSeverity.Error, "redaction-search-selection-stale",
                        "This search mark no longer matches the document. Search again before applying it.",
                        pageNumber: area.PageNumber));
                }
            } else {
                AddTextMatches(area, logical, matches);
            }
            if (includeHiddenOptionalContentText && area.PageNumber <= readDocument.Pages.Count) {
                PdfReadPage residualPage = readDocument.Pages[area.PageNumber - 1];
                if (residualPage.HasUnsupportedOptionalContentViewUsageApplications()) {
                    if (inconclusiveOptionalContentPages.Add(area.PageNumber)) {
                        findings.Add(new PdfDiagnosticFinding(
                            PdfDiagnosticSeverity.Error,
                            "redaction-plan-optional-content-inspection-inconclusive",
                            "Residual redaction inspection cannot determine optional-content visibility because the PDF uses unsupported View usage applications.",
                            pageNumber: area.PageNumber));
                    }
                } else if (!area.RequiresGlyphRewrite) {
                    if (!hiddenTextSpansByPage.TryGetValue(area.PageNumber, out IReadOnlyList<PdfTextSpan>? hiddenSpans)) {
                        hiddenSpans = residualPage.GetHiddenOptionalContentTextSpans(includeArtifactText: true);
                        hiddenTextSpansByPage.Add(area.PageNumber, hiddenSpans);
                    }
                    AddHiddenTextMatches(area, hiddenSpans, matches);
                }
            }
            if (area.IsHiddenOnly) {
                // A hidden-only selection paints nothing, so it has no claim on what shares its rectangle.
                AddAnnotationMatches(area, info.Pages, matches, readDocument);
                continue;
            }
            if (includeHiddenOptionalContentText && area.PageNumber <= readDocument.Pages.Count) {
                if (!hiddenImagePlacementsByPage.TryGetValue(area.PageNumber, out IReadOnlyList<PdfImagePlacement>? hiddenImagePlacements)) {
                    hiddenImagePlacements = readDocument.Pages[area.PageNumber - 1]
                        .GetImagePlacementsIncludingHiddenOptionalContent(area.PageNumber);
                    hiddenImagePlacementsByPage.Add(area.PageNumber, hiddenImagePlacements);
                }
                AddImagePlacementMatches(area, hiddenImagePlacements, matches, findings);
            } else {
                AddImageMatches(area, logical.Images, matches, findings);
            }
            AddAnnotationMatches(area, info.Pages, matches, readDocument);
            if (area.PageNumber <= readDocument.Pages.Count) {
                PdfReadPage page = readDocument.Pages[area.PageNumber - 1];
                if (!nestedPathPrimitivesByPage.TryGetValue(area.PageNumber, out IReadOnlyList<PdfPageVisualPrimitive>? primitives)) {
                    primitives = page.GetIdentityVisualPrimitives(cancellationToken);
                    nestedPathPrimitivesByPage.Add(area.PageNumber, primitives);
                }
                if (!generatedMarksByPage.TryGetValue(area.PageNumber, out HashSet<int>? generatedMarks)) {
                    generatedMarks = excludeGeneratedRedactionMarks
                        ? FindGeneratedRedactionMarks(areaArray, area.PageNumber, page, primitives) : new HashSet<int>();
                    generatedMarksByPage.Add(area.PageNumber, generatedMarks);
                }
                AddPathMatches(area, page, primitives, matches, findings, generatedMarks);
            }
        }

        var imageFingerprints = new Dictionary<int, IReadOnlyList<string>>();
        foreach (PdfRedactionMatch match in matches) {
            if (match.Kind != PdfRedactionMatchKind.ImagePlacement || match.ObjectNumber is not int imageObject) continue;
            if (!imageFingerprints.TryGetValue(imageObject, out IReadOnlyList<string>? fingerprints)) {
                fingerprints = readDocument.Objects.TryGetValue(imageObject, out PdfIndirectObject? indirect) && indirect.Value is PdfStream image
                    ? PdfRedactionApplier.FingerprintImageData(readDocument.Objects, image) : Array.Empty<string>();
                imageFingerprints.Add(imageObject, fingerprints);
            }
            match.ImageDataSha256 = fingerprints;
        }

        findings.Add(new PdfDiagnosticFinding(
            PdfDiagnosticSeverity.Info,
            "redaction-plan-only",
            "This plan reports intersections with the exact reviewed geometry. It does not remove or rewrite PDF content."));
        cancellationToken.ThrowIfCancellationRequested();

        return new PdfRedactionPlan(
            preflight,
            areaArray,
            matches.AsReadOnly(),
            findings.AsReadOnly(),
            searchCriteria: null,
            PdfRedactionPlan.ComputeSourceSha256(pdf),
            pageIdentities,
            reviewedTextObjectScopes);
    }

    /// <summary>Plans redaction impact for a PDF file.</summary>
    public static PdfRedactionPlan Plan(string path, IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null, CancellationToken cancellationToken = default) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        byte[] pdf = PdfDocumentSource.FromPath(path, options, cancellationToken).Bytes;
        return Plan(pdf, areas, layoutOptions, options, cancellationToken);
    }

    /// <summary>Plans redaction impact for a readable PDF stream.</summary>
    public static PdfRedactionPlan Plan(Stream stream, IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null, CancellationToken cancellationToken = default) {
        byte[] pdf = PdfDocumentSource.FromRemainingStream(stream, options, cancellationToken).Bytes;
        return Plan(pdf, areas, layoutOptions, options, cancellationToken);
    }

    private static void AddTextMatches(PdfRedactionArea area, PdfDocumentReadResult document, List<PdfRedactionMatch> matches) {
        foreach (PdfLogicalTextBlock block in document.TextBlocks) {
            if (block.PageNumber != area.PageNumber) {
                continue;
            }

            if (block.Spans.Count > 0 && !block.Spans.Any(span =>
                PdfTextSpanGeometry.IntersectsAreaAtCharacterLevel(span, area))) {
                continue;
            }

            PdfTextSpanBounds bounds = GetTextBlockBounds(block, document.Pages[block.PageNumber - 1]);
            if (!area.IntersectsRectangle(bounds.Left, bounds.Bottom, bounds.Width, bounds.Height)) {
                continue;
            }

            matches.Add(new PdfRedactionMatch(
                PdfRedactionMatchKind.TextBlock,
                area,
                block.PageNumber,
                bounds.Left,
                bounds.Bottom,
                bounds.Width,
                bounds.Height,
                block.Text,
                null,
                null));
        }
    }

    // Reports what the mark removes, not everything its rectangle touches. Returns false when a selected glyph is missing.
    private static bool AddSelectedGlyphMatches(PdfRedactionArea area, PdfRedactionGlyphSelection selection, IReadOnlyList<PdfTextSpan> spans,
        List<PdfRedactionMatch> matches, CancellationToken cancellationToken) {
        var resolved = new HashSet<int>();
        var blockText = new System.Text.StringBuilder();
        PdfTextSpanBounds? blockBounds = null;
        foreach (PdfTextSpan span in spans) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!area.TargetsSpan(PdfTextSpanGeometry.GetRedactionBounds(span)) || !span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> glyphs)) continue;
            PdfTextGlyph[] selected = glyphs.Where(glyph => selection.Contains(glyph.X, glyph.Y, glyph.EncodedBytes)).ToArray();
            if (selected.Length == 0) continue;
            PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries);
            double left = double.PositiveInfinity, bottom = double.PositiveInfinity;
            double right = double.NegativeInfinity, top = double.NegativeInfinity;
            foreach (PdfTextGlyph glyph in selected) {
                selection.MarkResolved(glyph.X, glyph.Y, glyph.EncodedBytes, resolved);
                PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetRedactionGlyphBounds(span, boundaries[glyph.TextStart], glyph.PaintedAdvance);
                left = Math.Min(left, bounds.Left); bottom = Math.Min(bottom, bounds.Bottom);
                right = Math.Max(right, bounds.Right); top = Math.Max(top, bounds.Top);
            }
            string text = string.Concat(selected.Select(static glyph => glyph.Text));
            if (selection.IsCompleteBlock) {
                blockText.Append(text);
                blockBounds = blockBounds is PdfTextSpanBounds current
                    ? new PdfTextSpanBounds(Math.Min(current.Left, left), Math.Min(current.Bottom, bottom), Math.Max(current.Right, right), Math.Max(current.Top, top))
                    : new PdfTextSpanBounds(left, bottom, right, top);
                continue;
            }
            matches.Add(new PdfRedactionMatch(PdfRedactionMatchKind.TextBlock, area, area.PageNumber, left, bottom, right - left, top - bottom, text, null, null));
        }
        // A block is one reviewed item, whatever text runs it was written in.
        if (blockBounds is PdfTextSpanBounds block) {
            matches.Add(new PdfRedactionMatch(PdfRedactionMatchKind.TextBlock, area, area.PageNumber, block.Left, block.Bottom,
                block.Right - block.Left, block.Top - block.Bottom, blockText.ToString(), null, null));
        }
        return resolved.Count == selection.Count;
    }

    private static double GetEffectiveFontSize(PdfLogicalTextBlock block) {
        return block.FontSize > 0D && !double.IsNaN(block.FontSize) && !double.IsInfinity(block.FontSize)
            ? block.FontSize
            : DefaultTextHeight;
    }

    private static PdfTextSpanBounds GetTextBlockBounds(PdfLogicalTextBlock block, PdfLogicalPage page) {
        if (block.Spans.Count > 0) {
            double left = double.MaxValue;
            double bottom = double.MaxValue;
            double right = double.MinValue;
            double top = double.MinValue;
            for (int index = 0; index < block.Spans.Count; index++) {
                PdfTextSpanBounds spanBounds = PdfTextSpanGeometry.GetRedactionBounds(block.Spans[index]);
                left = Math.Min(left, spanBounds.Left);
                bottom = Math.Min(bottom, spanBounds.Bottom);
                right = Math.Max(right, spanBounds.Right);
                top = Math.Max(top, spanBounds.Top);
            }

            return new PdfTextSpanBounds(left, bottom, right, top);
        }

        if (block.VisualBounds is PdfLogicalVisualBounds visualBounds) {
            PdfVisualBounds userBounds = page.TransformVisualBoundsToUser(
                visualBounds.Left,
                visualBounds.Top,
                visualBounds.Right,
                visualBounds.Bottom);
            return new PdfTextSpanBounds(userBounds.Left, userBounds.Top, userBounds.Right, userBounds.Bottom);
        }

        double fontSize = GetEffectiveFontSize(block);
        double x = Math.Min(block.XStart, block.XEnd);
        double width = Math.Max(1D, Math.Abs(block.XEnd - block.XStart));
        return new PdfTextSpanBounds(x, block.BaselineY - fontSize * 0.25D, x + width, block.BaselineY + fontSize);
    }

    private static void AddAnnotationMatches(PdfRedactionArea area, IReadOnlyList<PdfPageInfo> pages, List<PdfRedactionMatch> matches, PdfReadDocument document) {
        foreach (PdfPageInfo page in pages) {
            if (page.PageNumber != area.PageNumber) {
                continue;
            }

            foreach (PdfAnnotation annotation in page.Annotations) {
                if (!area.IntersectsUnderlay(annotation.X1, annotation.Y1, annotation.Width, annotation.Height) &&
                    !PdfRedactionApplier.CarriesSearchedText(document.Objects, document.Pages[page.PageNumber - 1].ObjectNumber, annotation, new[] { area })) {
                    continue;
                }

                matches.Add(new PdfRedactionMatch(
                    PdfRedactionMatchKind.Annotation,
                    area,
                    page.PageNumber,
                    annotation.X1,
                    annotation.Y1,
                    annotation.Width,
                    annotation.Height,
                    annotation.Contents,
                    annotation.Subtype,
                    annotation.ObjectNumber));
            }
        }
    }

    private static void AddImageMatches(PdfRedactionArea area, IReadOnlyList<PdfLogicalImage> images, List<PdfRedactionMatch> matches, List<PdfDiagnosticFinding> findings) {
        foreach (PdfLogicalImage image in images) {
            if (image.PageNumber != area.PageNumber) {
                continue;
            }

            foreach (PdfImagePlacement placement in image.Placements) {
                AddImagePlacementMatch(area, placement, matches, findings);
            }
        }
    }

    private static void AddImagePlacementMatches(
        PdfRedactionArea area,
        IReadOnlyList<PdfImagePlacement> placements,
        List<PdfRedactionMatch> matches,
        List<PdfDiagnosticFinding> findings) {
        for (int i = 0; i < placements.Count; i++) {
            PdfImagePlacement placement = placements[i];
            if (placement.PageNumber == area.PageNumber) AddImagePlacementMatch(area, placement, matches, findings);
        }
    }

    private static void AddImagePlacementMatch(
        PdfRedactionArea area,
        PdfImagePlacement placement,
        List<PdfRedactionMatch> matches,
        List<PdfDiagnosticFinding> findings) {
        if (!area.IntersectsRectangle(placement.X, placement.Y, placement.Width, placement.Height)) return;

        matches.Add(new PdfRedactionMatch(
            PdfRedactionMatchKind.ImagePlacement,
            area,
            placement.PageNumber,
            placement.X,
            placement.Y,
            placement.Width,
            placement.Height,
            null,
            null,
            placement.ObjectNumber == 0 ? null : placement.ObjectNumber,
            placement.ResourceName,
            placement));

        findings.Add(new PdfDiagnosticFinding(
            area.ContentScope == PdfRedactionContentScope.TextOnly ? PdfDiagnosticSeverity.Info : PdfDiagnosticSeverity.Warning,
            "redaction-plan-image-intersection",
            area.ContentScope == PdfRedactionContentScope.TextOnly
                ? "Redaction area intersects an image placement retained by the TextOnly policy."
                : "Redaction area intersects an image placement. Applying the plan rewrites supported image pixels and otherwise follows the configured fail-closed, whole-placement removal, or explicit visual-overlay policy.",
            placement.ObjectNumber == 0 ? null : placement.ObjectNumber,
            placement.PageNumber));
    }

    private static void AddHiddenTextMatches(
        PdfRedactionArea area,
        IReadOnlyList<PdfTextSpan> spans,
        List<PdfRedactionMatch> matches) {
        for (int i = 0; i < spans.Count; i++) {
            PdfTextSpan span = spans[i];
            PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetRedactionBounds(span);
            if (!area.IntersectsRectangle(bounds.Left, bounds.Bottom, bounds.Width, bounds.Height)) continue;
            matches.Add(new PdfRedactionMatch(
                PdfRedactionMatchKind.TextBlock,
                area,
                area.PageNumber,
                bounds.Left,
                bounds.Bottom,
                bounds.Width,
                bounds.Height,
                span.Text,
                null,
                null));
        }
    }

    private static void AddPathMatches(
        PdfRedactionArea area,
        PdfReadPage page,
        IReadOnlyList<PdfPageVisualPrimitive> primitives,
        List<PdfRedactionMatch> matches,
        List<PdfDiagnosticFinding> findings,
        HashSet<int> generatedMarks) {
        for (int i = 0; i < primitives.Count; i++) {
            PdfPageVisualPrimitive primitive = primitives[i];
            if (generatedMarks.Contains(i)) continue;
            double strokePadding = primitive.HasStrokePaint ? Math.Max(0D, primitive.StrokeWidth) / 2D : 0D;
            double left = primitive.X - strokePadding;
            double top = primitive.Y - strokePadding;
            double right = primitive.X + primitive.Width + strokePadding;
            double bottom = primitive.Y + primitive.Height + strokePadding;
            PdfVisualBounds userBounds = page.TransformVisualBoundsToUser(left, top, right, bottom);
            if (!area.IntersectsUnderlay(userBounds.Left, userBounds.Top, userBounds.Width, userBounds.Height)) continue;
            matches.Add(new PdfRedactionMatch(
                PdfRedactionMatchKind.VectorPath,
                area,
                area.PageNumber,
                userBounds.Left,
                userBounds.Top,
                userBounds.Width,
                userBounds.Height,
                null,
                primitive.Kind.ToString(),
                null));
            findings.Add(new PdfDiagnosticFinding(
                area.ContentScope == PdfRedactionContentScope.TextOnly ? PdfDiagnosticSeverity.Info : PdfDiagnosticSeverity.Warning,
                primitive.ContentOrderKey?.Depth > 1
                    ? "redaction-plan-nested-vector-intersection"
                    : "redaction-plan-vector-intersection",
                area.ContentScope == PdfRedactionContentScope.TextOnly
                    ? "The redaction area intersects vector content retained by the TextOnly policy."
                    : primitive.ContentOrderKey?.Depth > 1
                    ? "The redaction area intersects vector content inside a nested Form XObject. The current writer cannot remove that path; applied-plan verification will remain unverified while the nested vector content remains."
                    : "The redaction area intersects a page-level vector path. Applying the plan removes supported intersecting paths when path removal is enabled; otherwise applied-plan verification remains unverified while the path survives.",
                pageNumber: area.PageNumber));
        }
    }

    // Marks on adjacent lines overlap, so one area's mark must not read as residual content in the next.
    private static HashSet<int> FindGeneratedRedactionMarks(
        IReadOnlyList<PdfRedactionArea> areas,
        int pageNumber,
        PdfReadPage page,
        IReadOnlyList<PdfPageVisualPrimitive> primitives) {
        var marks = new HashSet<int>();
        foreach (PdfRedactionArea area in areas) {
            if (area.PageNumber != pageNumber) continue;
            int mark = FindGeneratedRedactionMarkIndex(area, page.TransformBoundsToVisual(area.X, area.Y, area.Right, area.Top), primitives);
            if (mark >= 0) marks.Add(mark);
        }
        return marks;
    }

    private static int FindGeneratedRedactionMarkIndex(
        PdfRedactionArea area,
        PdfVisualBounds visualArea,
        IReadOnlyList<PdfPageVisualPrimitive> primitives) {
        for (int i = primitives.Count - 1; i >= 0; i--) {
            PdfPageVisualPrimitive primitive = primitives[i];
            if (primitive.ContentOrderKey?.Depth > 1) {
                continue;
            }
            bool rectangleMark = primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle && primitive.HasFillPaint && !primitive.HasStrokePaint;
            bool exactMark = area.ExactGeometry is not null && (primitive.HasFillPaint || primitive.HasStrokePaint);
            // Freehand review bounds include the round-cap stroke radius;
            // primitive bounds describe the centerline before that padding.
            double strokePadding = area.ExactGeometry?.Kind == PdfRedactionRegionKind.Freehand && primitive.HasStrokePaint
                ? Math.Max(0D, primitive.StrokeWidth) / 2D : 0D;
            if ((rectangleMark || exactMark) && AreClose(primitive.X - strokePadding, visualArea.Left) &&
                AreClose(primitive.Y - strokePadding, visualArea.Top) &&
                AreClose(primitive.Width + strokePadding * 2D, visualArea.Width) &&
                AreClose(primitive.Height + strokePadding * 2D, visualArea.Height)) {
                return i;
            }
        }
        return -1;
    }

    private static bool AreClose(double left, double right) => Math.Abs(left - right) < 0.0001D;

    private static bool Intersects(double ax, double ay, double aw, double ah, double bx, double by, double bw, double bh) {
        return ax < bx + bw &&
            ax + aw > bx &&
            ay < by + bh &&
            ay + ah > by;
    }
}
