using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<HtmlRenderVisual> SliceBlockVisuals(HtmlRenderFlowBlock block, double start, double end) {
        return SliceVisuals(block.Visuals, start, end, includeEndAnchor: end >= block.PagedPaintExtent - 0.0001D);
    }

    private IReadOnlyList<HtmlRenderVisual> SliceVisuals(IEnumerable<HtmlRenderVisual> sourceVisuals, double start, double end, bool includeEndAnchor = false) {
        var fragment = new List<HtmlRenderVisual>();
        foreach (HtmlRenderVisual visual in sourceVisuals) {
            int firstFragment = fragment.Count;
            try {
                double visualTop = visual.LayoutY;
                if (visual is HtmlRenderLogicalTextGroup { IsFlowAnchor: true }) {
                    // Interior points belong to the following fragment; the
                    // terminal point belongs to the final fragment exactly once.
                    if (visualTop >= start - 0.0001D
                        && (visualTop < end - 0.0001D || includeEndAnchor && visualTop <= end + 0.0001D)) {
                        fragment.Add(visual.Translate(0D, -start, fragment.Count));
                    }
                    continue;
                }
                bool containsFlowAnchor = _pageFloatEntries.Count > 0 && ContainsFlowAnchor(visual);
                double visualBottom = visual.LayoutY + visual.LayoutHeight;
                // Paint-neutral wrappers can retain the CSS box height while
                // their visible children continue across printed pages.
                if (visual is HtmlRenderSemanticGroup semanticOverflow) {
                    visualBottom = Math.Max(visualBottom, MaximumScrollBottom(semanticOverflow.Visuals));
                } else if (visual is HtmlRenderLayoutRegion regionOverflow) {
                    visualBottom = Math.Max(visualBottom, MaximumScrollBottom(regionOverflow.Visuals));
                } else if (visual is HtmlRenderEffectGroup effectOverflow) {
                    visualBottom = Math.Max(visualBottom, MaximumScrollBottom(new[] { effectOverflow }));
                } else if (visual is HtmlRenderClipGroup { ClipVertical: false } visibleVerticalClip) {
                    visualBottom = Math.Max(visualBottom, MaximumScrollBottom(visibleVerticalClip.Visuals));
                }
                double intersectionTop = Math.Max(start, visualTop);
                double intersectionBottom = Math.Min(end, visualBottom);
                if (intersectionBottom <= intersectionTop + 0.0001D && !containsFlowAnchor) continue;

                bool fullyContained = visualTop >= start - 0.0001D && visualBottom <= end + 0.0001D;
                if (fullyContained && !containsFlowAnchor) {
                    fragment.Add(visual.Translate(0D, -start, fragment.Count));
                    continue;
                }

                if (visual is HtmlRenderClipGroup clipGroup) {
                    IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(clipGroup.Visuals, start, end, includeEndAnchor);
                    if (children.Count > 0) {
                        fragment.Add(new HtmlRenderClipGroup(
                            clipGroup.ClipX,
                            clipGroup.ClipY - start,
                            clipGroup.ClipWidth,
                            clipGroup.ClipHeight,
                            clipGroup.ClipHorizontal,
                            clipGroup.ClipVertical,
                            children,
                            fragment.Count,
                            clipGroup.Source,
                            Math.Max(start, clipGroup.LayoutY) - start,
                            clipGroup.IsViewportOverflow));
                    }
                    continue;
                }

                if (visual is HtmlRenderSemanticGroup semanticGroup) {
                    IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(semanticGroup.Visuals, start, end, includeEndAnchor);
                    if (children.Count > 0) {
                        fragment.Add(new HtmlRenderSemanticGroup(
                            semanticGroup.Role,
                            semanticGroup.X,
                            semanticGroup.Y - start,
                            semanticGroup.Width,
                            Math.Max(0.01D, intersectionBottom - intersectionTop),
                            children,
                            fragment.Count,
                            semanticGroup.Source,
                            semanticGroup.ColumnSpan,
                            semanticGroup.RowSpan,
                            semanticGroup.HeaderScope,
                            semanticGroup.LayoutY - start,
                            semanticGroup.StructureElementKey));
                    }
                    continue;
                }

                if (visual is HtmlRenderLayoutRegion layoutRegion) {
                    IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(layoutRegion.Visuals, start, end, includeEndAnchor);
                    fragment.Add(new HtmlRenderLayoutRegion(
                        layoutRegion.SourceKey,
                        layoutRegion.RegionKind,
                        layoutRegion.SourceText,
                        layoutRegion.Position,
                        layoutRegion.FloatSide,
                        layoutRegion.ZIndex,
                        layoutRegion.BackgroundLayerCount,
                        layoutRegion.BoxShadowLayerCount,
                        layoutRegion.BackgroundColor,
                        layoutRegion.X,
                        layoutRegion.Y - start,
                        layoutRegion.Width,
                        Math.Max(0.01D, intersectionBottom - intersectionTop),
                        children,
                        fragment.Count,
                        layoutRegion.Source,
                        layoutRegion.LayoutY - start));
                    continue;
                }

                if (visual is HtmlRenderBookmarkAnchor bookmarkAnchor) {
                    if (bookmarkAnchor.LayoutY >= start - 0.0001D && bookmarkAnchor.LayoutY < end - 0.0001D) {
                        fragment.Add(new HtmlRenderBookmarkAnchor(
                            bookmarkAnchor.SemanticNodeId,
                            bookmarkAnchor.Text,
                            bookmarkAnchor.X,
                            bookmarkAnchor.Y - start,
                            bookmarkAnchor.Width,
                            bookmarkAnchor.Height,
                            fragment.Count,
                            bookmarkAnchor.Source,
                            bookmarkAnchor.LayoutY - start));
                    }
                    continue;
                }

                if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
                    IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(logicalTextGroup.Visuals, start, end, includeEndAnchor);
                    if (children.Count > 0) {
                        fragment.Add(new HtmlRenderLogicalTextGroup(
                            ResolveLogicalText(children, logicalTextGroup.Text),
                            logicalTextGroup.X,
                            logicalTextGroup.Y - start,
                            logicalTextGroup.Width,
                            Math.Max(0.01D, intersectionBottom - intersectionTop),
                            children,
                            fragment.Count,
                            logicalTextGroup.Source,
                            logicalTextGroup.LayoutY - start,
                            logicalScope: logicalTextGroup.LogicalScope));
                    }
                    continue;
                }

                if (visual is HtmlRenderEffectGroup effectGroup) {
                    bool translated = TryGetVerticalPaintTranslation(effectGroup.Transform, out double verticalTranslation);
                    double childStart = translated ? start - verticalTranslation : start;
                    double childEnd = translated ? end - verticalTranslation : end;
                    IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(effectGroup.Visuals, childStart, childEnd, includeEndAnchor);
                    if (children.Count > 0) {
                        if (translated && Math.Abs(verticalTranslation) > 0.0001D) {
                            // Child fragments are rebased to their pre-transform window;
                            // move them back to this page before applying the effect.
                            children = children.Select((child, index) =>
                                child.Translate(0D, -verticalTranslation, index)).ToList();
                        }
                        double translatedY = -start;
                        OfficeTransform transform = OfficeTransform.Translate(0D, -translatedY)
                            .Then(effectGroup.Transform)
                            .Then(OfficeTransform.Translate(0D, translatedY));
                        fragment.Add(new HtmlRenderEffectGroup(
                            effectGroup.X,
                            effectGroup.Y - start,
                            effectGroup.Width,
                            Math.Max(0.01D, intersectionBottom - intersectionTop),
                            transform,
                            effectGroup.Opacity,
                            children,
                            fragment.Count,
                            effectGroup.Source,
                            Math.Max(start, effectGroup.LayoutY) - start));
                    }
                    continue;
                }

                if (visual is HtmlRenderLayoutBox layoutBox) {
                    fragment.Add(new HtmlRenderLayoutBox(layoutBox.X,
                        layoutBox.Y + intersectionTop - layoutBox.LayoutY - start,
                        layoutBox.Width, Math.Max(0.01D, intersectionBottom - intersectionTop),
                        fragment.Count, layoutBox.Source, intersectionTop - start,
                        layoutBox.IsAbsolutePrintOverflow));
                    continue;
                }

                if (visual is HtmlRenderPathClipGroup pathClip && containsFlowAnchor) {
                    // Atomic paint clipping keeps the complete child tree on
                    // each page. Source points must instead belong only to
                    // their fragment; retain the same path and outer paint clip.
                    IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(pathClip.Visuals, start, end, includeEndAnchor);
                    if (children.Count > 0) {
                        var slicedPath = new HtmlRenderPathClipGroup(pathClip.ClipX, pathClip.ClipY - start,
                            pathClip.ClipPath, children, 0, pathClip.Source, pathClip.LayoutY - start);
                        double clipY = intersectionTop - start;
                        fragment.Add(new HtmlRenderClipGroup(pathClip.X, clipY, pathClip.Width,
                            Math.Max(0.01D, intersectionBottom - intersectionTop),
                            clipHorizontal: false, clipVertical: true, new[] { slicedPath },
                            fragment.Count, pathClip.Source, clipY));
                    }
                    continue;
                }

                if (visual is HtmlRenderImage
                    || visual is HtmlRenderDrawing
                    || visual is HtmlRenderImagePattern
                    || visual is HtmlRenderPathClipGroup
                    || visual is HtmlRenderShape
                    || visual is HtmlRenderAnchorFragment) {
                    fragment.Add(CreateVerticallyClippedVisualFragment(visual, start, intersectionTop, intersectionBottom, fragment.Count));
                    continue;
                }

                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.VisualFragmentUnsupported, "A visual crossing a forced page boundary could not be represented safely in the current fragment.", HtmlDiagnosticSeverity.Warning, visual.Source, visual.Kind.ToString());
            } finally {
                if (visual.StackingContext != null) {
                    for (int index = firstFragment; index < fragment.Count; index++) {
                        visual.CopyStackingContextTo(fragment[index]);
                    }
                }
            }
        }

        return fragment;
    }

    private bool ContainsFlowAnchor(HtmlRenderVisual visual) {
        foreach (HtmlRenderVisual child in EnumeratePageFloatVisuals(new[] { visual })) {
            ChargeLayoutOperation("page-float fragment anchors");
            if (child is HtmlRenderLogicalTextGroup { IsFlowAnchor: true }) return true;
        }
        return false;
    }

    private static HtmlRenderClipGroup CreateVerticallyClippedVisualFragment(
        HtmlRenderVisual visual,
        double fragmentStart,
        double intersectionTop,
        double intersectionBottom,
        int paintOrder) {
        double clipY = intersectionTop - fragmentStart;
        return new HtmlRenderClipGroup(
            visual.X,
            clipY,
            visual.Width,
            Math.Max(0.01D, intersectionBottom - intersectionTop),
            clipHorizontal: false,
            clipVertical: true,
            new[] { visual.Translate(0D, -fragmentStart, 0) },
            paintOrder,
            visual.Source,
            clipY);
    }

}
