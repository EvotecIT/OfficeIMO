using System.Globalization;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Finds reviewable field candidates in static PDF page content.</summary>
internal static class PdfStaticFormRecognizer {
    /// <summary>Analyzes page geometry and native text, with optional caller-supplied OCR labels, without changing the source PDF.</summary>
    internal static PdfStaticFormRecognitionReport Analyze(
        PdfDocument source,
        PdfStaticFormRecognitionOptions? options = null,
        IReadOnlyList<PdfStaticFormTextEvidence>? ocrText = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(source, nameof(source));
        PdfStaticFormRecognitionOptions effective = options ?? new PdfStaticFormRecognitionOptions();
        effective.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        var snapshot = source.GetReadSnapshot(cancellationToken: cancellationToken);
        byte[] pdf = snapshot.Bytes;
        PdfReadDocument document = snapshot.Document;
        int[] pageNumbers = effective.PageSelection?.ToPageNumbers(document.Pages.Count, nameof(effective.PageSelection))
            ?? Enumerable.Range(1, document.Pages.Count).ToArray();
        pageNumbers = pageNumbers.Distinct().ToArray();
        if (pageNumbers.Length > effective.MaxPages) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, effective.MaxPages, pageNumbers.Length);
        }
        IReadOnlyList<PdfStaticFormTextEvidence> suppliedText = ocrText ?? Array.Empty<PdfStaticFormTextEvidence>();
        if (suppliedText.Count > effective.MaxOcrTextItems) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, effective.MaxOcrTextItems, suppliedText.Count);
        }
        foreach (PdfStaticFormTextEvidence item in suppliedText) {
            Guard.NotNull(item, nameof(ocrText));
            if (item.PageNumber > document.Pages.Count) throw new ArgumentOutOfRangeException(nameof(ocrText), "OCR evidence names a page outside the PDF.");
        }

        PdfDocumentReadResult logical = PdfDocumentReadEngine.Read(document, new PdfReadOptions {
            Profile = PdfReadProfile.Fast,
            PageSelection = PdfPageSelection.From(pageNumbers),
            Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = effective.MaxPages }
        }, cancellationToken);
        var proposed = new List<Candidate>();
        var diagnostics = new List<PdfStaticFormRecognitionDiagnostic>();
        var pageDirections = new Dictionary<int, PdfReadingDirection>();
        long candidateScanWork = 0;
        foreach (int pageNumber in pageNumbers) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfLogicalPage page = logical.PagesBySourcePageNumber[pageNumber][0];
            (double pageWidth, double pageHeight) = page.GetVisualPageSize();
            PdfReadPage readPage = document.Pages[pageNumber - 1];
            IReadOnlyList<PdfPageVisualPrimitive> primitives = readPage.GetIdentityVisualPrimitives(
                scaleStrokeWidthWithTransform: true, cancellationToken: cancellationToken);
            IReadOnlyList<PdfPageDrawingEffectTransition> effects = readPage.GetIdentityGraphicsEffectTransitions(cancellationToken);
            int imagePlacementCount = page.Images.Sum(static image => image.Placements.Count);
            var filledAreas = new List<PaintArea>();
            foreach (PdfPageVisualPrimitive painted in primitives) {
                cancellationToken.ThrowIfCancellationRequested();
                if (painted.HasFillPaint && painted.FillOpacity != 0D && painted.Width > 0D && painted.Height > 0D) {
                    PdfPageDrawingEffect paintEffect = PdfReadPage.ResolveDrawingEffect(effects, painted.PaintOrder,
                        contentOrderKey: painted.ContentOrderKey);
                    bool normalBlend = paintEffect.BlendMode == OfficeBlendMode.Normal &&
                        paintEffect.SoftMask is null && !paintEffect.HasUnresolvedSoftMask;
                    var visible = new VisualRect(painted.X, painted.Y, painted.X + painted.Width, painted.Y + painted.Height);
                    if (painted.ClipPath is PdfPageClipPath clip && clip.IsRectangle && clip.IsExact && !clip.ContainsTextClipping) {
                        visible = new VisualRect(Math.Max(visible.Left, clip.X), Math.Max(visible.Top, clip.Y),
                            Math.Min(visible.Right, clip.X + clip.Width), Math.Min(visible.Bottom, clip.Y + clip.Height));
                        if (visible.Area == 0D) continue;
                    }
                    filledAreas.Add(new PaintArea(visible, painted.PaintOrder, painted.ContentOrderKey,
                        normalBlend && IsEmptyFill(painted), normalBlend && IsOpaqueWhiteFill(painted),
                        normalBlend && IsOpaqueCover(painted),
                        normalBlend && IsOpaqueCover(painted) ? painted.FillColor : null));
                }
            }
            List<Label> labels = GetLabels(page, suppliedText, pageWidth, pageHeight, filledAreas, effects,
                out List<VisualRect> nativeTextBounds, ref candidateScanWork, effective.MaxCandidateScanWork, cancellationToken);
            pageDirections[pageNumber] = PdfTextDirectionAnalysis.Resolve(PdfReadingDirection.Auto,
                labels.Select(static label => label.Text));
            for (int candidateIndex = 0; candidateIndex < primitives.Count; candidateIndex++) {
                PdfPageVisualPrimitive primitive = primitives[candidateIndex];
                cancellationToken.ThrowIfCancellationRequested();
                if (!TryGetCandidate(primitive, pageWidth, pageHeight, out VisualRect visual, out PdfStaticFormEvidenceKind evidence)) continue;
                candidateScanWork = checked(candidateScanWork +
                    ((long)primitives.Count + imagePlacementCount) * (filledAreas.Count + 1L) +
                    effects.Count + labels.Count * 2L + nativeTextBounds.Count + proposed.Count + page.FormWidgets.Count +
                    page.Annotations.Count + page.LinkAnnotations.Count);
                if (candidateScanWork > effective.MaxCandidateScanWork) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts,
                        effective.MaxCandidateScanWork, candidateScanWork);
                }
                PdfPageDrawingEffect effect = PdfReadPage.ResolveDrawingEffect(effects, primitive.PaintOrder,
                    contentOrderKey: primitive.ContentOrderKey);
                if (effect.SoftMask is not null || effect.HasUnresolvedSoftMask ||
                    effect.BlendMode != OfficeBlendMode.Normal) {
                    AddDiagnostic("unsupported-effect", pageNumber,
                        "A visual field candidate uses compositing that cannot prove an empty field.");
                    continue;
                }
                if (primitive.StrokeColor is OfficeColor strokeColor &&
                    primitive.StrokeGradient is null && primitive.StrokeRadialGradient is null &&
                    primitive.StrokeTilingPattern is null) {
                    bool visible = IsOpaqueWhiteFill(primitive) && primitive.FillColor is OfficeColor ownFill
                        ? ColorsContrast(strokeColor, ownFill)
                        : filledAreas.Any(area => IsCandidateBackdrop(area, primitive, visual, evidence)) ||
                          HasContrastingBackdrop(filledAreas, OutlinePaintBounds(primitive, visual), strokeColor,
                              primitive.PaintOrder, primitive.ContentOrderKey, outlinedStroke: true);
                    if (!visible) {
                        AddDiagnostic("invisible-outline", pageNumber,
                            "A field outline cannot be distinguished from its painted background.");
                        continue;
                    }
                }
                if (IsCoveredByLaterOpaqueFill(filledAreas, OutlinePaintBounds(primitive, visual),
                    primitive.PaintOrder, primitive.ContentOrderKey)) {
                    AddDiagnostic("occluded-outline", pageNumber,
                        "A visual field outline is covered by later opaque paint.");
                    continue;
                }
                if (HasPaintedInterior(filledAreas, primitive, visual, evidence, cancellationToken)) continue;
                if (HasInteriorMark(primitives, filledAreas, effects, candidateIndex, visual, evidence, cancellationToken) ||
                    HasImageInterior(page, filledAreas, visual, cancellationToken)) {
                    AddDiagnostic("occupied-field", pageNumber, "A visual field candidate contains a painted mark.");
                    continue;
                }
                if (OverlapsExistingWidget(page, visual)) {
                    AddDiagnostic("existing-widget", pageNumber, "A visual candidate overlaps an existing form widget.");
                    continue;
                }
                if (OverlapsExistingAnnotation(page, visual)) {
                    AddDiagnostic("existing-annotation", pageNumber, "A visual candidate overlaps an existing annotation.");
                    continue;
                }
                if (nativeTextBounds.Any(bounds => OverlapArea(bounds, visual) > Math.Min(bounds.Area, visual.Area) * 0.05D) ||
                    labels.Any(label => label.IsOcr && OverlapArea(label.Bounds, visual) > Math.Min(label.Bounds.Area, visual.Area) * 0.05D)) continue;
                Label? labelMatch = FindLabel(labels, visual, evidence, pageDirections[pageNumber]);
                if (labelMatch is null) continue;
                double confidence = evidence switch {
                    PdfStaticFormEvidenceKind.OutlinedField => 0.84D,
                    PdfStaticFormEvidenceKind.CheckBox => 0.79D,
                    _ => 0.72D
                };
                confidence *= labelMatch.Confidence;
                if (labelMatch.IsOcr) confidence *= 0.9D;
                if (confidence < effective.MinimumConfidence) {
                    AddDiagnostic("low-confidence", pageNumber, "A visual field candidate was below the selected confidence threshold.");
                    continue;
                }
                if (proposed.Any(candidate => candidate.PageNumber == pageNumber &&
                    OverlapArea(candidate.Visual, visual) > Math.Min(candidate.Visual.Area, visual.Area) * 0.6D)) {
                    AddDiagnostic("overlapping-proposals", pageNumber, "Overlapping visual field candidates need manual review.");
                    continue;
                }
                proposed.Add(new Candidate(pageNumber, visual, evidence, labelMatch, confidence));
                if (proposed.Count > effective.MaxProposals) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.FormFields, effective.MaxProposals, proposed.Count);
                }
            }
        }

        var usedNames = new HashSet<string>(StringComparer.Ordinal);
        foreach (PdfFormField field in document.FormFields) {
            if (string.IsNullOrWhiteSpace(field.Name)) continue;
            string name = field.Name!;
            usedNames.Add(name);
            for (int separator = name.IndexOf('.'); separator >= 0; separator = name.IndexOf('.', separator + 1)) {
                usedNames.Add(name.Substring(0, separator));
            }
        }
        var proposals = new List<PdfStaticFormFieldProposal>(proposed.Count);
        int currentPage = 0;
        int tabIndex = 0;
        foreach (Candidate candidate in proposed.OrderBy(static candidate => candidate.PageNumber)
                     .ThenBy(static candidate => candidate.Visual.Top)
                     .ThenBy(candidate => pageDirections[candidate.PageNumber] == PdfReadingDirection.RightToLeft
                         ? -candidate.Visual.Left : candidate.Visual.Left)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (candidate.PageNumber != currentPage) { currentPage = candidate.PageNumber; tabIndex = 0; }
            PdfLogicalPage page = logical.PagesBySourcePageNumber[candidate.PageNumber][0];
            PdfPageRectangle rectangle = page.MapVisualRectangleToUserSpace(candidate.Visual.Left, candidate.Visual.Top, candidate.Visual.Right, candidate.Visual.Bottom);
            string name = UniqueName(candidate.Label.Text, usedNames);
            proposals.Add(new PdfStaticFormFieldProposal(
                proposals.Count, candidate.PageNumber, ++tabIndex, candidate.Label.Text, name,
                candidate.Evidence == PdfStaticFormEvidenceKind.CheckBox ? PdfFormFieldCreationKind.CheckBox : PdfFormFieldCreationKind.Text,
                rectangle,
                new PdfLogicalVisualBounds(candidate.Visual.Left, candidate.Visual.Top, candidate.Visual.Right, candidate.Visual.Bottom),
                candidate.Confidence, candidate.Evidence, candidate.Label.IsOcr));
        }
        return new PdfStaticFormRecognitionReport(pdf, snapshot.Options, proposals, diagnostics);

        void AddDiagnostic(string code, int pageNumber, string message) {
            if (diagnostics.Count < effective.MaxDiagnostics) {
                diagnostics.Add(new PdfStaticFormRecognitionDiagnostic(code, pageNumber, message));
            } else if (diagnostics[diagnostics.Count - 1].Code != "diagnostics-truncated") {
                diagnostics[diagnostics.Count - 1] = new PdfStaticFormRecognitionDiagnostic("diagnostics-truncated", pageNumber,
                    "Additional recognition diagnostics were omitted at the configured MaxDiagnostics limit.");
            }
        }
    }

    private static List<Label> GetLabels(PdfLogicalPage page, IReadOnlyList<PdfStaticFormTextEvidence> ocrText,
        double pageWidth, double pageHeight,
        IReadOnlyList<PaintArea> filledAreas, IReadOnlyList<PdfPageDrawingEffectTransition> effects,
        out List<VisualRect> nativeTextBounds, ref long candidateScanWork, int maxCandidateScanWork,
        CancellationToken cancellationToken) {
        var labels = new List<Label>();
        nativeTextBounds = new List<VisualRect>();
        foreach (PdfLogicalTextBlock block in page.TextBlocks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (block.XEnd <= block.XStart) continue;
            PdfLogicalVisualBounds bounds = block.VisualBounds ?? ToVisualBounds(page, block);
            if (!Valid(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom, pageWidth, pageHeight)) continue;
            var visual = new VisualRect(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom);
            bool visibleText = block.Spans.Count == 0;
            bool provableLabel = block.Spans.Count == 0;
            string text = NormalizeLabel(block.Text);
            if (block.Spans.Count > 0) {
                bool uncertainEffect = false;
                var visibleSpans = new List<PdfTextSpan>(block.Spans.Count);
                VisualRect? labelBounds = null;
                foreach (PdfTextSpan span in block.Spans) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!span.IsVisible || (span.Color?.A ?? 255) <= 3) continue;
                    PdfTextSpanBounds spanBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(span);
                    if (span.ClipPath is PdfPageClipPath clip) {
                        var paintedText = PdfPageClipPath.Rectangle(spanBounds.Left,
                            page.Height - spanBounds.Top, spanBounds.Right - spanBounds.Left,
                            spanBounds.Top - spanBounds.Bottom);
                        if (clip.Width <= 0D || clip.Height <= 0D ||
                            clip.CanProveNoPositiveAreaIntersection(paintedText)) continue;
                    }
                    PdfVisualBounds spanProjected = page.TransformBoundsToVisual(spanBounds.Left, spanBounds.Bottom,
                        spanBounds.Right, spanBounds.Top);
                    if (!Valid(spanProjected.Left, spanProjected.Top, spanProjected.Right, spanProjected.Bottom,
                        pageWidth, pageHeight)) continue;
                    var spanVisual = new VisualRect(spanProjected.Left, spanProjected.Top,
                        spanProjected.Right, spanProjected.Bottom);
                    PdfPageDrawingEffect effect = PdfReadPage.ResolveDrawingEffect(effects, span.PaintOrder,
                        contentOrderKey: span.ContentOrderKey);
                    if (effect.SoftMask is not null || effect.HasUnresolvedSoftMask ||
                        effect.BlendMode != OfficeBlendMode.Normal) uncertainEffect = true;
                    bool covered = false;
                    foreach (PaintArea area in filledAreas) {
                        cancellationToken.ThrowIfCancellationRequested();
                        candidateScanWork++;
                        if (candidateScanWork > maxCandidateScanWork) {
                            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts,
                                maxCandidateScanWork, candidateScanWork);
                        }
                        if (IsLaterCover(area, spanVisual, span.PaintOrder, span.ContentOrderKey)) {
                            covered = true;
                            break;
                        }
                    }
                    if (!covered) {
                        foreach (PdfLogicalImage image in page.Images) {
                            foreach (PdfImagePlacement placement in image.Placements) {
                                cancellationToken.ThrowIfCancellationRequested();
                                candidateScanWork++;
                                if (candidateScanWork > maxCandidateScanWork) {
                                    throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts,
                                        maxCandidateScanWork, candidateScanWork);
                                }
                                if (IsLaterOpaqueImageCover(page, image, placement, spanVisual,
                                    span.PaintOrder, span.ContentOrderKey)) { covered = true; break; }
                            }
                            if (covered) break;
                        }
                    }
                    if (covered) continue;
                    if (span.Color is OfficeColor ink && ink.R >= 245 && ink.G >= 245 && ink.B >= 245 &&
                        !HasContrastingBackdrop(filledAreas, spanVisual, span.PaintOrder, span.ContentOrderKey)) {
                        // Keep the text as possible field occupancy, but do not infer a label without visible contrast.
                        nativeTextBounds.Add(spanVisual);
                        uncertainEffect = true;
                        continue;
                    }
                    visibleSpans.Add(span);
                    nativeTextBounds.Add(spanVisual);
                    labelBounds = labelBounds.HasValue
                        ? new VisualRect(Math.Min(labelBounds.Value.Left, spanVisual.Left),
                            Math.Min(labelBounds.Value.Top, spanVisual.Top),
                            Math.Max(labelBounds.Value.Right, spanVisual.Right),
                            Math.Max(labelBounds.Value.Bottom, spanVisual.Bottom))
                        : spanVisual;
                }
                visibleText = visibleSpans.Count > 0;
                provableLabel = visibleText && !uncertainEffect;
                if (labelBounds.HasValue) visual = labelBounds.Value;
                if (visibleSpans.Count != block.Spans.Count) {
                    text = NormalizeLabel(string.Concat(visibleSpans.Select(static span => span.Text)));
                }
            }
            if (visibleText && block.Spans.Count == 0) nativeTextBounds.Add(visual);
            if (!provableLabel || text.Length == 0 || text.Length > 80) continue;
            labels.Add(new Label(text, visual, block.Confidence, isOcr: false));
        }
        foreach (PdfStaticFormTextEvidence item in ocrText) {
            cancellationToken.ThrowIfCancellationRequested();
            if (item.PageNumber != page.PageNumber) continue;
            string text = NormalizeLabel(item.Text);
            if (text.Length == 0 || text.Length > 80 || !Valid(item.Left, item.Top, item.Right, item.Bottom, pageWidth, pageHeight)) continue;
            var bounds = new VisualRect(item.Left, item.Top, item.Right, item.Bottom);
            if (labels.Any(label => string.Equals(label.Text, text, StringComparison.OrdinalIgnoreCase) &&
                OverlapArea(label.Bounds, bounds) >= Math.Max(label.Bounds.Area, bounds.Area) * 0.8D)) continue;
            if (item.Confidence >= 0.8D) {
                // A high-confidence caller correction supersedes conflicting native extraction at the same location.
                labels.RemoveAll(label => !label.IsOcr &&
                    OverlapArea(label.Bounds, bounds) > Math.Min(label.Bounds.Area, bounds.Area) * 0.5D);
            }
            labels.Add(new Label(text, bounds, item.Confidence, isOcr: true));
        }
        return labels;
    }

    private static PdfLogicalVisualBounds ToVisualBounds(PdfLogicalPage page, PdfLogicalTextBlock block) {
        double size = Math.Max(1D, block.FontSize);
        PdfVisualBounds visual = page.TransformBoundsToVisual(block.XStart, block.BaselineY - size * 0.25D, block.XEnd, block.BaselineY + size * 0.85D);
        return new PdfLogicalVisualBounds(visual.Left, visual.Top, visual.Right, visual.Bottom);
    }

    private static bool TryGetCandidate(PdfPageVisualPrimitive primitive, double pageWidth, double pageHeight, out VisualRect bounds, out PdfStaticFormEvidenceKind evidence) {
        bounds = default;
        evidence = default;
        if (!primitive.HasStrokePaint || primitive.StrokeOpacity == 0D) return false;
        if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
            if (primitive.Width >= 9D && primitive.Width <= 22D && primitive.Height >= 9D && primitive.Height <= 22D &&
                Math.Abs(primitive.Width - primitive.Height) <= 3D && IsEmptyFill(primitive)) {
                bounds = new VisualRect(primitive.X, primitive.Y, primitive.X + primitive.Width, primitive.Y + primitive.Height);
                evidence = PdfStaticFormEvidenceKind.CheckBox;
            } else if (primitive.Width >= 45D && primitive.Width <= 500D && primitive.Height >= 13D && primitive.Height <= 45D && IsEmptyFill(primitive)) {
                bounds = new VisualRect(primitive.X, primitive.Y, primitive.X + primitive.Width, primitive.Y + primitive.Height);
                evidence = PdfStaticFormEvidenceKind.OutlinedField;
            }
        } else if (primitive.Kind == PdfPageVisualPrimitiveKind.Line &&
                   Math.Abs(primitive.Y1 - primitive.Y2) <= 1D &&
                   Math.Abs(primitive.X2 - primitive.X1) >= 60D && Math.Abs(primitive.X2 - primitive.X1) <= 500D) {
            double left = Math.Min(primitive.X1, primitive.X2);
            bounds = new VisualRect(left, primitive.Y1 - 18D, Math.Max(primitive.X1, primitive.X2), primitive.Y1 + 2D);
            evidence = PdfStaticFormEvidenceKind.Underline;
        }
        if (primitive.ClipPath is PdfPageClipPath clip &&
            (!clip.IsRectangle || !clip.IsExact || clip.ContainsTextClipping ||
             clip.X > bounds.Left || clip.Y > bounds.Top ||
             clip.X + clip.Width < bounds.Right || clip.Y + clip.Height < bounds.Bottom)) return false;
        return bounds.Area > 0D && Valid(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom, pageWidth, pageHeight);
    }

    private static bool IsEmptyFill(PdfPageVisualPrimitive primitive) =>
        primitive.FillOpacity == 0D || !primitive.HasFillPaint ||
        primitive.FillGradient is null && primitive.FillRadialGradient is null && primitive.FillTilingPattern is null &&
        primitive.FillColor is OfficeIMO.Drawing.OfficeColor color && color.R >= 245 && color.G >= 245 && color.B >= 245;

    private static bool IsOpaqueWhiteFill(PdfPageVisualPrimitive primitive) =>
        IsOpaqueCover(primitive) &&
        primitive.FillColor is OfficeColor color && color.R >= 245 && color.G >= 245 && color.B >= 245;

    private static bool IsOpaqueCover(PdfPageVisualPrimitive primitive) =>
        primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle && primitive.HasFillPaint &&
        (primitive.FillOpacity ?? 1D) >= 0.999D &&
        primitive.FillGradient is null && primitive.FillRadialGradient is null && primitive.FillTilingPattern is null &&
        (primitive.ClipPath is not PdfPageClipPath clip ||
         clip.IsRectangle && clip.IsExact && !clip.ContainsTextClipping) &&
        primitive.FillColor is OfficeColor color && color.A >= 254;

    private static bool HasPaintedInterior(
        IReadOnlyList<PaintArea> filledAreas,
        PdfPageVisualPrimitive outline, VisualRect candidate, PdfStaticFormEvidenceKind evidence,
        CancellationToken cancellationToken) {
        PaintArea? latestPaint = null;
        bool painted = false;
        foreach (PaintArea area in filledAreas) {
            cancellationToken.ThrowIfCancellationRequested();
            if (latestPaint.HasValue && !IsLater(area.PaintOrder, area.ContentOrderKey,
                latestPaint.Value.PaintOrder, latestPaint.Value.ContentOrderKey)) continue;
            if (IsCandidateBackdrop(area, outline, candidate, evidence)) {
                latestPaint = area;
                painted = false;
                continue;
            }
            if (area.IsEmpty) {
                if (area.IsOpaqueWhite && area.Bounds.Left <= candidate.Left && area.Bounds.Top <= candidate.Top &&
                    area.Bounds.Right >= candidate.Right && area.Bounds.Bottom >= candidate.Bottom) {
                    latestPaint = area;
                    painted = false;
                }
            } else {
                if (area.Bounds.Area > candidate.Area * 1.25D ||
                    OverlapArea(area.Bounds, candidate) < candidate.Area * 0.8D) continue;
                latestPaint = area;
                painted = true;
            }
        }
        return painted;
    }

    private static bool IsCandidateBackdrop(PaintArea area, PdfPageVisualPrimitive outline,
        VisualRect candidate, PdfStaticFormEvidenceKind evidence) =>
        evidence == PdfStaticFormEvidenceKind.OutlinedField && area.IsOpaqueCover &&
        area.Color is OfficeColor background && outline.StrokeColor is OfficeColor ink &&
        ColorsContrast(ink, background) &&
        IsLater(outline.PaintOrder, outline.ContentOrderKey, area.PaintOrder, area.ContentOrderKey) &&
        area.Bounds.Left <= candidate.Left && area.Bounds.Top <= candidate.Top &&
        area.Bounds.Right >= candidate.Right && area.Bounds.Bottom >= candidate.Bottom &&
        area.Bounds.Area <= candidate.Area * 1.25D;

    private static bool HasInteriorMark(IReadOnlyList<PdfPageVisualPrimitive> primitives,
        IReadOnlyList<PaintArea> filledAreas, IReadOnlyList<PdfPageDrawingEffectTransition> effects,
        int candidateIndex, VisualRect candidate, PdfStaticFormEvidenceKind evidence,
        CancellationToken cancellationToken) {
        const double inset = 0.2D;
        var interior = new VisualRect(candidate.Left + inset, candidate.Top + inset,
            candidate.Right - inset, candidate.Bottom - inset);
        PdfPageVisualPrimitive candidatePrimitive = primitives[candidateIndex];
        for (int index = 0; index < primitives.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (index == candidateIndex) continue;
            PdfPageVisualPrimitive primitive = primitives[index];
            if (filledAreas.Any(area =>
                !IsLater(area.PaintOrder, area.ContentOrderKey, primitive.PaintOrder, primitive.ContentOrderKey) &&
                !IsLater(primitive.PaintOrder, primitive.ContentOrderKey, area.PaintOrder, area.ContentOrderKey) &&
                IsCandidateBackdrop(area, candidatePrimitive, candidate, evidence))) continue;
            bool paintedAfterCandidate = IsLater(primitive.PaintOrder, primitive.ContentOrderKey,
                candidatePrimitive.PaintOrder, candidatePrimitive.ContentOrderKey);
            bool stroked = primitive.HasStrokePaint && primitive.StrokeOpacity != 0D;
            bool hasFillPaint = primitive.HasFillPaint && primitive.FillOpacity != 0D;
            if (!stroked && !hasFillPaint) continue;
            double left = primitive.Kind == PdfPageVisualPrimitiveKind.Line ? Math.Min(primitive.X1, primitive.X2) : primitive.X;
            double top = primitive.Kind == PdfPageVisualPrimitiveKind.Line ? Math.Min(primitive.Y1, primitive.Y2) : primitive.Y;
            double right = primitive.Kind == PdfPageVisualPrimitiveKind.Line ? Math.Max(primitive.X1, primitive.X2) : primitive.X + primitive.Width;
            double bottom = primitive.Kind == PdfPageVisualPrimitiveKind.Line ? Math.Max(primitive.Y1, primitive.Y2) : primitive.Y + primitive.Height;
            if (stroked) {
                // The parser stores the RMS transform scale. Its largest singular scale can be
                // sqrt(2) times larger, so use that conservative envelope for occupancy.
                double strokePadding = Math.Max(0D, primitive.StrokeWidth) * Math.Sqrt(2D) / 2D;
                left -= strokePadding;
                top -= strokePadding;
                right += strokePadding;
                bottom += strokePadding;
            }
            if (primitive.ClipPath is PdfPageClipPath clip && clip.IsRectangle && clip.IsExact &&
                !clip.ContainsTextClipping) {
                left = Math.Max(left, clip.X);
                top = Math.Max(top, clip.Y);
                right = Math.Min(right, clip.X + clip.Width);
                bottom = Math.Min(bottom, clip.Y + clip.Height);
            }
            if (right <= left || bottom <= top) continue;
            PdfPageDrawingEffect effect = PdfReadPage.ResolveDrawingEffect(effects, primitive.PaintOrder,
                contentOrderKey: primitive.ContentOrderKey);
            bool uncertainEffect = effect.BlendMode != OfficeBlendMode.Normal || effect.SoftMask is not null ||
                effect.HasUnresolvedSoftMask;
            bool filled = hasFillPaint && (uncertainEffect || !IsEmptyFill(primitive) ||
                HasContrastingBackdrop(filledAreas, new VisualRect(left, top, right, bottom),
                    primitive.PaintOrder, primitive.ContentOrderKey));
            if (!stroked && !filled) continue;
            if (primitive.Kind != PdfPageVisualPrimitiveKind.Line &&
                (right - left) * (bottom - top) > candidate.Area * 1.25D &&
                !paintedAfterCandidate && left <= candidate.Left && top <= candidate.Top &&
                right >= candidate.Right && bottom >= candidate.Bottom) continue;
            if (!filled && (right - left < 1D || bottom - top < 1D)) continue;
            var visible = new VisualRect(Math.Max(left, interior.Left), Math.Max(top, interior.Top),
                Math.Min(right, interior.Right), Math.Min(bottom, interior.Bottom));
            if (visible.Area > (filled ? 0D : 0.5D) &&
                !IsCoveredByLaterOpaqueFill(filledAreas, visible, primitive.PaintOrder,
                    primitive.ContentOrderKey)) return true;
        }
        return false;
    }

    private static bool HasImageInterior(PdfLogicalPage page,
        IReadOnlyList<PaintArea> filledAreas,
        VisualRect candidate, CancellationToken cancellationToken) {
        var interior = new VisualRect(candidate.Left + 0.2D, candidate.Top + 0.2D,
            candidate.Right - 0.2D, candidate.Bottom - 0.2D);
        foreach (PdfLogicalImage image in page.Images) {
            foreach (PdfImagePlacement placement in image.Placements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (placement.IsHiddenOptionalContent || placement.Opacity <= 0D ||
                    placement.Width <= 0D || placement.Height <= 0D) continue;
                PdfVisualBounds mapped = page.TransformBoundsToVisual(placement.X, placement.Y,
                    placement.X + placement.Width, placement.Y + placement.Height);
                var visible = new VisualRect(mapped.Left, mapped.Top, mapped.Right, mapped.Bottom);
                if (placement.Clip is { IsRectangle: true, IsExact: true, ContainsTextClipping: false } clip) {
                    PdfVisualBounds clipped = page.TransformBoundsToVisual(clip.X,
                        page.Height - clip.Y - clip.Height, clip.X + clip.Width, page.Height - clip.Y);
                    visible = new VisualRect(Math.Max(visible.Left, clipped.Left), Math.Max(visible.Top, clipped.Top),
                        Math.Min(visible.Right, clipped.Right), Math.Min(visible.Bottom, clipped.Bottom));
                }
                var inside = new VisualRect(Math.Max(visible.Left, interior.Left), Math.Max(visible.Top, interior.Top),
                    Math.Min(visible.Right, interior.Right), Math.Min(visible.Bottom, interior.Bottom));
                if (inside.Area > 0.5D &&
                    !IsCoveredByLaterOpaqueFill(filledAreas, inside, placement.PaintOrder,
                        placement.ContentOrderKey)) return true;
            }
        }
        return false;
    }

    private static bool IsCoveredByLaterOpaqueFill(
        IReadOnlyList<PaintArea> filledAreas,
        VisualRect painted, double paintOrder, PdfContentOrderKey? contentOrderKey) =>
        filledAreas.Any(area => IsLaterCover(area, painted, paintOrder, contentOrderKey));

    private static bool IsLaterCover(PaintArea area, VisualRect painted,
        double paintOrder, PdfContentOrderKey? contentOrderKey) =>
        area.IsOpaqueCover &&
            IsLater(area.PaintOrder, area.ContentOrderKey, paintOrder, contentOrderKey) &&
            area.Bounds.Left <= painted.Left && area.Bounds.Top <= painted.Top &&
            area.Bounds.Right >= painted.Right && area.Bounds.Bottom >= painted.Bottom;

    private static bool IsLaterOpaqueImageCover(PdfLogicalPage page, PdfLogicalImage image,
        PdfImagePlacement placement, VisualRect painted, double paintOrder, PdfContentOrderKey? contentOrderKey) {
        if (!IsLater(placement.PaintOrder, placement.ContentOrderKey, paintOrder, contentOrderKey) ||
            placement.IsHiddenOptionalContent || placement.Width <= 0D || placement.Height <= 0D ||
            placement.Opacity < 0.999D || placement.HasSoftMask || placement.HasUnsupportedBlendMode ||
            placement.HasUnsupportedPaintState || placement.HasUnsupportedImagePaintEffect ||
            placement.EffectiveBlendMode != OfficeBlendMode.Normal || image.SourceImage.IsImageMask ||
            image.SourceImage.HasTransparencyMask ||
            !(Math.Abs(placement.B) < 0.000001D && Math.Abs(placement.C) < 0.000001D ||
              Math.Abs(placement.A) < 0.000001D && Math.Abs(placement.D) < 0.000001D) ||
            placement.Clip is { IsRectangle: false } or { IsExact: false } or { ContainsTextClipping: true }) return false;
        PdfVisualBounds mapped = page.TransformBoundsToVisual(placement.X, placement.Y,
            placement.X + placement.Width, placement.Y + placement.Height);
        var visible = new VisualRect(mapped.Left, mapped.Top, mapped.Right, mapped.Bottom);
        if (placement.Clip is { } clip) {
            PdfVisualBounds clipped = page.TransformBoundsToVisual(clip.X,
                page.Height - clip.Y - clip.Height, clip.X + clip.Width, page.Height - clip.Y);
            visible = new VisualRect(Math.Max(visible.Left, clipped.Left), Math.Max(visible.Top, clipped.Top),
                Math.Min(visible.Right, clipped.Right), Math.Min(visible.Bottom, clipped.Bottom));
        }
        return visible.Left <= painted.Left && visible.Top <= painted.Top &&
            visible.Right >= painted.Right && visible.Bottom >= painted.Bottom;
    }

    private static bool HasContrastingBackdrop(IReadOnlyList<PaintArea> filledAreas, VisualRect outline,
        double paintOrder, PdfContentOrderKey? contentOrderKey) =>
        HasContrastingBackdrop(filledAreas, outline, OfficeColor.White, paintOrder, contentOrderKey);

    private static bool HasContrastingBackdrop(IReadOnlyList<PaintArea> filledAreas, VisualRect outline,
        OfficeColor ink, double paintOrder, PdfContentOrderKey? contentOrderKey, bool outlinedStroke = false) {
        PaintArea? latestFullCover = null;
        foreach (PaintArea area in filledAreas) {
            if (!area.IsOpaqueCover || !IsLater(paintOrder, contentOrderKey, area.PaintOrder, area.ContentOrderKey) ||
                OverlapArea(area.Bounds, outline) <= 0D) continue;
            if (area.Bounds.Left <= outline.Left && area.Bounds.Top <= outline.Top &&
                area.Bounds.Right >= outline.Right && area.Bounds.Bottom >= outline.Bottom &&
                (!latestFullCover.HasValue || IsLater(area.PaintOrder, area.ContentOrderKey,
                    latestFullCover.Value.PaintOrder, latestFullCover.Value.ContentOrderKey))) latestFullCover = area;
        }
        OfficeColor backdrop = latestFullCover?.Color ?? OfficeColor.White;
        if (!ColorsContrast(ink, backdrop)) return false;
        foreach (PaintArea area in filledAreas) {
            if (!IsLater(paintOrder, contentOrderKey, area.PaintOrder, area.ContentOrderKey) ||
                latestFullCover.HasValue && !IsLater(area.PaintOrder, area.ContentOrderKey,
                    latestFullCover.Value.PaintOrder, latestFullCover.Value.ContentOrderKey) ||
                OverlapArea(area.Bounds, outline) <= 0D) continue;
            // Paint wholly inside a stroked outline does not cover its visible border.
            if (outlinedStroke && area.Bounds.Left > outline.Left && area.Bounds.Top > outline.Top &&
                area.Bounds.Right < outline.Right && area.Bounds.Bottom < outline.Bottom) continue;
            if (!area.IsOpaqueCover || area.Color is not OfficeColor color || !ColorsContrast(ink, color)) return false;
        }
        return true;
    }

    private static bool ColorsContrast(OfficeColor ink, OfficeColor backdrop) =>
        Math.Max(Math.Abs(ink.R - backdrop.R),
            Math.Max(Math.Abs(ink.G - backdrop.G), Math.Abs(ink.B - backdrop.B))) >= 45;

    internal static bool IsLater(double candidateOrder, PdfContentOrderKey? candidateKey,
        double earlierOrder, PdfContentOrderKey? earlierKey) =>
        candidateKey != null && earlierKey != null
            ? candidateKey.CompareTo(earlierKey) > 0
            : candidateOrder > earlierOrder;

    private static VisualRect OutlinePaintBounds(PdfPageVisualPrimitive primitive, VisualRect candidate) {
        double padding = Math.Max(0D, primitive.StrokeWidth) / 2D;
        if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
            return new VisualRect(Math.Min(primitive.X1, primitive.X2) - padding,
                Math.Min(primitive.Y1, primitive.Y2) - padding,
                Math.Max(primitive.X1, primitive.X2) + padding,
                Math.Max(primitive.Y1, primitive.Y2) + padding);
        }
        return new VisualRect(candidate.Left - padding, candidate.Top - padding,
            candidate.Right + padding, candidate.Bottom + padding);
    }

    private static Label? FindLabel(IReadOnlyList<Label> labels, VisualRect field,
        PdfStaticFormEvidenceKind evidence, PdfReadingDirection direction) {
        Label? best = null;
        double bestDistance = double.MaxValue;
        foreach (Label label in labels) {
            VisualRect bounds = label.Bounds;
            double centerDifference = Math.Abs((bounds.Top + bounds.Bottom) / 2D - (field.Top + field.Bottom) / 2D);
            double distance = double.MaxValue;
            if (centerDifference <= Math.Max(10D, field.Height * 0.65D)) {
                if (bounds.Right <= field.Left && field.Left - bounds.Right <= 120D) distance = field.Left - bounds.Right;
                if ((evidence == PdfStaticFormEvidenceKind.CheckBox || direction == PdfReadingDirection.RightToLeft) &&
                    bounds.Left >= field.Right && bounds.Left - field.Right <= 120D) {
                    distance = Math.Min(distance, bounds.Left - field.Right);
                }
            }
            if (bounds.Bottom <= field.Top && field.Top - bounds.Bottom <= 30D &&
                bounds.Left <= field.Right && bounds.Right >= field.Left - 15D) {
                distance = Math.Min(distance, 15D + field.Top - bounds.Bottom);
            }
            if (distance < bestDistance) { best = label; bestDistance = distance; }
        }
        return best;
    }

    private static bool OverlapsExistingWidget(PdfLogicalPage page, VisualRect bounds) {
        foreach (PdfLogicalFormWidget widget in page.FormWidgets) {
            if (widget.X2 <= widget.X1 || widget.Y2 <= widget.Y1) continue;
            PdfSelectionQuad visual = page.MapUserSpaceRectangleToVisual(widget.X1, widget.Y1, widget.X2, widget.Y2);
            var widgetBounds = new VisualRect(visual.Left, visual.Top, visual.Right, visual.Bottom);
            if (OverlapArea(bounds, widgetBounds) > Math.Min(bounds.Area, widgetBounds.Area) * 0.05D) return true;
        }
        return false;
    }

    private static bool OverlapsExistingAnnotation(PdfLogicalPage page, VisualRect bounds) {
        foreach (PdfAnnotation annotation in page.Annotations) {
            if (!annotation.HasReadableRectangle || annotation.X2 <= annotation.X1 || annotation.Y2 <= annotation.Y1) continue;
            PdfSelectionQuad visual = page.MapUserSpaceRectangleToVisual(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2);
            var annotationBounds = new VisualRect(visual.Left, visual.Top, visual.Right, visual.Bottom);
            if (OverlapArea(bounds, annotationBounds) > Math.Min(bounds.Area, annotationBounds.Area) * 0.05D) return true;
        }
        foreach (PdfLinkAnnotation link in page.LinkAnnotations) {
            if (link.X2 <= link.X1 || link.Y2 <= link.Y1) continue;
            PdfSelectionQuad visual = page.MapUserSpaceRectangleToVisual(link.X1, link.Y1, link.X2, link.Y2);
            var linkBounds = new VisualRect(visual.Left, visual.Top, visual.Right, visual.Bottom);
            if (OverlapArea(bounds, linkBounds) > Math.Min(bounds.Area, linkBounds.Area) * 0.05D) return true;
        }
        return false;
    }

    private static string UniqueName(string label, HashSet<string> used) {
        var builder = new StringBuilder(label.Length);
        foreach (char character in label) {
            if (char.IsLetterOrDigit(character)) builder.Append(char.ToLowerInvariant(character));
            else if (builder.Length > 0 && builder[builder.Length - 1] != '_') builder.Append('_');
        }
        string stem = builder.ToString().Trim('_');
        if (stem.Length > 50) stem = stem.Substring(0, 50).TrimEnd('_');
        if (stem.Length == 0) stem = "field";
        string candidate = stem;
        int suffix = 2;
        while (!used.Add(candidate)) {
            candidate = stem + "_" + suffix.ToString(CultureInfo.InvariantCulture);
            suffix++;
        }
        return candidate;
    }

    private static string NormalizeLabel(string text) => text.Trim().TrimEnd(':', '：').Trim();

    private static bool Valid(double left, double top, double right, double bottom, double width, double height) =>
        IsFinite(left) && IsFinite(top) && IsFinite(right) && IsFinite(bottom) && left >= 0D && top >= 0D && right <= width && bottom <= height && right > left && bottom > top;
    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private static double OverlapArea(VisualRect first, VisualRect second) =>
        Math.Max(0D, Math.Min(first.Right, second.Right) - Math.Max(first.Left, second.Left)) *
        Math.Max(0D, Math.Min(first.Bottom, second.Bottom) - Math.Max(first.Top, second.Top));

    private readonly struct VisualRect {
        internal VisualRect(double left, double top, double right, double bottom) { Left = left; Top = top; Right = right; Bottom = bottom; }
        internal double Left { get; }
        internal double Top { get; }
        internal double Right { get; }
        internal double Bottom { get; }
        internal double Height => Bottom - Top;
        internal double Area => Math.Max(0D, Right - Left) * Math.Max(0D, Height);
    }

    private readonly struct PaintArea {
        internal PaintArea(VisualRect bounds, double paintOrder, PdfContentOrderKey? contentOrderKey,
            bool isEmpty, bool isOpaqueWhite, bool isOpaqueCover, OfficeColor? color) {
            Bounds = bounds;
            PaintOrder = paintOrder;
            ContentOrderKey = contentOrderKey;
            IsEmpty = isEmpty;
            IsOpaqueWhite = isOpaqueWhite;
            IsOpaqueCover = isOpaqueCover;
            Color = color;
        }
        internal VisualRect Bounds { get; }
        internal double PaintOrder { get; }
        internal PdfContentOrderKey? ContentOrderKey { get; }
        internal bool IsEmpty { get; }
        internal bool IsOpaqueWhite { get; }
        internal bool IsOpaqueCover { get; }
        internal OfficeColor? Color { get; }
    }

    private sealed class Label {
        internal Label(string text, VisualRect bounds, double confidence, bool isOcr) { Text = text; Bounds = bounds; Confidence = confidence; IsOcr = isOcr; }
        internal string Text { get; }
        internal VisualRect Bounds { get; }
        internal double Confidence { get; }
        internal bool IsOcr { get; }
    }

    private sealed class Candidate {
        internal Candidate(int pageNumber, VisualRect visual, PdfStaticFormEvidenceKind evidence, Label label, double confidence) {
            PageNumber = pageNumber; Visual = visual; Evidence = evidence; Label = label; Confidence = confidence;
        }
        internal int PageNumber { get; }
        internal VisualRect Visual { get; }
        internal PdfStaticFormEvidenceKind Evidence { get; }
        internal Label Label { get; }
        internal double Confidence { get; }
    }
}
