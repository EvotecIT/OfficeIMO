namespace OfficeIMO.Html;

/// <summary>Creates bounded, cancellable page projections while preserving only intersecting scene nodes.</summary>
internal sealed class HtmlRenderVisualProjector {
    private readonly int _maximumVisuals;
    private readonly CancellationToken _cancellationToken;
    private readonly HashSet<HtmlRenderLogicalTextGroup> _ownedLogicalText = new();
    private readonly HashSet<HtmlRenderText> _ownedText = new();
    private readonly HashSet<HtmlRenderVisual> _ownedNavigation = new();
    private int _materializedVisuals;

    internal HtmlRenderVisualProjector(int maximumVisuals, CancellationToken cancellationToken) {
        if (maximumVisuals <= 0) throw new ArgumentOutOfRangeException(nameof(maximumVisuals));
        _maximumVisuals = maximumVisuals;
        _cancellationToken = cancellationToken;
    }

    internal IReadOnlyList<HtmlRenderVisual> Project(
        IEnumerable<HtmlRenderVisual> visuals,
        double sourceX,
        double sourceY,
        double width,
        double height,
        double outputX,
        double outputY) {
        if (visuals == null) throw new ArgumentNullException(nameof(visuals));
        var projected = new List<HtmlRenderVisual>();
        double right = sourceX + width;
        double bottom = sourceY + height;
        double offsetX = outputX - sourceX;
        double offsetY = outputY - sourceY;
        foreach (HtmlRenderVisual visual in visuals.OrderBy(item => item.PaintOrder)) {
            HtmlRenderVisual? result = Project(visual, sourceX, sourceY, right, bottom, offsetX, offsetY, projected.Count);
            if (result != null) projected.Add(result);
        }
        return projected.AsReadOnly();
    }

    private HtmlRenderVisual? Project(
        HtmlRenderVisual visual,
        double left,
        double top,
        double right,
        double bottom,
        double offsetX,
        double offsetY,
        int paintOrder) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (!Intersects(visual, left, top, right, bottom)) return null;

        if (visual is HtmlRenderClipGroup clipGroup) {
            return ProjectGroup(clipGroup.Visuals, children => clipGroup.ProjectPaint(children, offsetX, offsetY, paintOrder));
        }
        if (visual is HtmlRenderPathClipGroup pathClipGroup) {
            return ProjectGroup(pathClipGroup.Visuals, children => pathClipGroup.ProjectPaint(children, offsetX, offsetY, paintOrder));
        }
        if (visual is HtmlRenderEffectGroup effectGroup) {
            return ProjectGroup(effectGroup.Visuals, children => effectGroup.ProjectPaint(children, offsetX, offsetY, paintOrder));
        }
        if (visual is HtmlRenderSemanticGroup semanticGroup) {
            return ProjectGroup(semanticGroup.Visuals, children => semanticGroup.ProjectPaint(children, offsetX, offsetY, paintOrder));
        }
        if (visual is HtmlRenderLayoutRegion layoutRegion) {
            return ProjectGroup(layoutRegion.Visuals, children => layoutRegion.ProjectPaint(children, offsetX, offsetY, paintOrder));
        }
        if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
            bool ownsLogicalText = _ownedLogicalText.Add(logicalTextGroup);
            return ProjectGroup(logicalTextGroup.Visuals,
                children => logicalTextGroup.ProjectPaint(children, offsetX, offsetY, paintOrder, ownsLogicalText));
        }
        if (visual is HtmlRenderFormField formField) {
            return ProjectGroup(formField.Visuals, children => formField.ProjectPaint(children, offsetX, offsetY, paintOrder));
        }
        if ((visual is HtmlRenderNamedDestination || visual is HtmlRenderBookmarkAnchor)
            && !_ownedNavigation.Add(visual)) {
            return null;
        }
        if (visual is HtmlRenderText text && !_ownedText.Add(text)) {
            CountVisual();
            HtmlRenderVisual translated = text.TranslatePaint(offsetX, offsetY, 0);
            CountVisual();
            return new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.Artifact,
                translated.X, translated.Y, translated.Width, translated.Height,
                new[] { translated }, paintOrder, translated.Source, layoutY: translated.LayoutY);
        }

        CountVisual();
        return visual.TranslatePaint(offsetX, offsetY, paintOrder);

        HtmlRenderVisual? ProjectGroup(
            IEnumerable<HtmlRenderVisual> children,
            Func<IEnumerable<HtmlRenderVisual>, HtmlRenderVisual> create) {
            var projectedChildren = new List<HtmlRenderVisual>();
            foreach (HtmlRenderVisual child in children.OrderBy(item => item.PaintOrder)) {
                HtmlRenderVisual? projectedChild = Project(
                    child, left, top, right, bottom, offsetX, offsetY, projectedChildren.Count);
                if (projectedChild != null) projectedChildren.Add(projectedChild);
            }
            if (projectedChildren.Count == 0) return null;
            CountVisual();
            return create(projectedChildren);
        }
    }

    private void CountVisual() {
        _cancellationToken.ThrowIfCancellationRequested();
        if (++_materializedVisuals > _maximumVisuals) {
            throw new InvalidOperationException(
                $"HTML render projection exceeded MaxProjectedVisuals {_maximumVisuals}.");
        }
    }

    private static bool Intersects(HtmlRenderVisual visual, double left, double top, double right, double bottom) =>
        visual.X < right && visual.X + visual.Width > left
        && visual.Y < bottom && visual.Y + visual.Height > top;
}
