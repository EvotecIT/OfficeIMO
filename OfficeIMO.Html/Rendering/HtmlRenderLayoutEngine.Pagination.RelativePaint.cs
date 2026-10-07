using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool _hasRelativePagedPaint;

    /// <summary>
    /// Relative positioning changes paint, not fragmentation. Project its paint
    /// after normal-flow pagination so forced-break slack and page masters remain
    /// part of the physical coordinate system.
    /// </summary>
    private IReadOnlyList<HtmlRenderPage> ProjectRelativePagedPaint(List<HtmlRenderPage> pages) {
        if (!_hasRelativePagedPaint) return pages;
        var projection = new RelativePagePaintProjection(this, pages);
        return projection.Project();
    }

    private sealed class RelativePagePaintProjection {
        private const double Epsilon = 0.0001D;
        private readonly HtmlRenderLayoutEngine _engine;
        private readonly List<HtmlRenderPage> _pages;
        private readonly List<double> _starts = new() { 0D };
        private readonly HashSet<HtmlRenderVisual> _ownedText = new();
        private readonly HashSet<HtmlRenderVisual> _ownedNavigation = new();
        private readonly HashSet<HtmlRenderVisual> _relativeTrees = new();
        private int _materialized;

        internal RelativePagePaintProjection(HtmlRenderLayoutEngine engine, List<HtmlRenderPage> pages) {
            _engine = engine;
            _pages = pages;
            foreach (HtmlRenderPage page in pages) AppendWindow(page);
        }

        internal IReadOnlyList<HtmlRenderPage> Project() {
            int flowPageCount = _pages.Count;
            double maximum = _starts[_starts.Count - 1];
            for (int source = 0; source < flowPageCount; source++) {
                foreach (HtmlRenderVisual visual in _pages[source].Scene) {
                    Inspect(visual, source, double.NegativeInfinity, double.PositiveInfinity, OfficeTransform.Identity, ref maximum);
                }
            }
            while (maximum > _starts[_starts.Count - 1] + Epsilon) {
                _engine.ChargeLayoutOperation("relative paint overflow pages");
                string? name = _pages[_pages.Count - 1].PageName;
                HtmlCssPageGeometry geometry = _engine._pageRules.ResolveGeometry(_pages.Count + 1, name, _engine._options);
                _engine.SetActivePageGeometry(geometry);
                _engine.ValidateSurface(geometry.Width, geometry.Height);
                _engine.CommitPage(_pages, _engine.CreatePageVisuals(geometry.Width, geometry.Height, geometry), geometry, name);
                AppendWindow(_pages[_pages.Count - 1]);
            }

            var scenes = _pages.Select(page => new List<HtmlRenderVisual>()).ToList();
            for (int source = 0; source < flowPageCount; source++) {
                foreach (HtmlRenderVisual visual in _pages[source].Scene) {
                    if (!_relativeTrees.Contains(visual)) {
                        scenes[source].Add(visual);
                        continue;
                    }
                    foreach (KeyValuePair<int, HtmlRenderVisual> item in Route(
                        visual, source, double.NegativeInfinity, double.PositiveInfinity, OfficeTransform.Identity)) {
                        scenes[item.Key].Add(item.Value);
                    }
                }
            }
            for (int index = flowPageCount; index < _pages.Count; index++) scenes[index].InsertRange(0, _pages[index].Scene);
            return _pages.Select((page, index) => page.WithScene(scenes[index])).ToList().AsReadOnly();
        }

        private void AppendWindow(HtmlRenderPage page) =>
            _starts.Add(_starts[_starts.Count - 1] + Math.Max(1D, page.Height - page.Margins.Top - page.Margins.Bottom));

        private bool Inspect(HtmlRenderVisual visual, int source, double clipTop, double clipBottom,
            OfficeTransform transform, ref double maximum) {
            _engine.ChargeLayoutOperation("relative paint window inspection");
            if (visual.IsOutOfFlowPaint) return false;
            Constrain(visual, transform, ref clipTop, ref clipBottom);
            bool relative = false;
            if (visual.PaintChildren is { } children) {
                if (visual is HtmlRenderEffectGroup effect) transform = effect.Transform.Then(transform);
                // A singular effect has no two-dimensional paint support. Keep
                // its existing source representation; it cannot supply an
                // invertible physical-window clip or require overflow pages.
                if (!transform.TryInvert(out _)) return false;
                foreach (HtmlRenderVisual child in children) {
                    relative |= Inspect(child, source, clipTop, clipBottom, transform, ref maximum);
                }
            } else if (CanMove(visual)) {
                relative = true;
                HtmlRenderRectangle bounds = PaintBounds(visual, transform);
                double bottom = Math.Min(bounds.Bottom, clipBottom);
                if (bounds.Width > Epsilon && bottom > Math.Max(bounds.Y, clipTop) + Epsilon) {
                    maximum = Math.Max(maximum, _starts[source] + bottom - _pages[source].Margins.Top);
                }
            }
            if (relative) _relativeTrees.Add(visual);
            return relative;
        }

        private Dictionary<int, HtmlRenderVisual> Route(HtmlRenderVisual visual, int source, double clipTop, double clipBottom,
            OfficeTransform transform) {
            _engine.ChargeLayoutOperation("relative paint routing");
            if (!_relativeTrees.Contains(visual) || visual.IsOutOfFlowPaint
                || visual is HtmlRenderLogicalTextGroup { IsFlowAnchor: true }) return AtSource(visual, source);
            Constrain(visual, transform, ref clipTop, ref clipBottom);
            if (visual.PaintChildren is { } children) {
                OfficeTransform childTransform = visual is HtmlRenderEffectGroup effect ? effect.Transform.Then(transform) : transform;
                if (!childTransform.TryInvert(out _)) return AtSource(visual, source);
                var groups = new Dictionary<int, List<HtmlRenderVisual>>();
                foreach (HtmlRenderVisual child in children) {
                    foreach (KeyValuePair<int, HtmlRenderVisual> item in Route(child, source, clipTop, clipBottom, childTransform)) {
                        if (!groups.TryGetValue(item.Key, out List<HtmlRenderVisual>? group)) groups[item.Key] = group = new();
                        group.Add(item.Value);
                    }
                }
                var result = new Dictionary<int, HtmlRenderVisual>();
                foreach (KeyValuePair<int, List<HtmlRenderVisual>> item in groups.OrderBy(item => item.Key)) {
                    bool unchanged = item.Key == source && item.Value.Count == children.Count
                        && item.Value.Where((child, index) => !ReferenceEquals(child, children[index])).Any() == false
                        && visual is not HtmlRenderClipGroup { IsFlowFragment: true, RelativePaintOffsetY: not 0D };
                    if (unchanged) {
                        result[item.Key] = visual;
                    } else {
                        Count();
                        double dy = OffsetY(source, item.Key);
                        double dx = _pages[item.Key].Margins.Left - _pages[source].Margins.Left;
                        bool ownsText = visual is not HtmlRenderLogicalTextGroup || _ownedText.Add(visual);
                        result[item.Key] = visual.ProjectPaintChildren(item.Value, dx, dy, visual.PaintOrder, ownsText);
                    }
                }
                return result;
            }
            if (!CanMove(visual)) return AtSource(visual, source);

            HtmlRenderRectangle paintBounds = PaintBounds(visual, transform);
            if (paintBounds.Width <= Epsilon) return AtSource(visual, source);
            double top = Math.Max(paintBounds.Y, clipTop);
            double bottom = Math.Min(paintBounds.Bottom, clipBottom);
            var routed = new Dictionary<int, HtmlRenderVisual>();
            if (bottom <= top + Epsilon) return routed;
            double virtualTop = _starts[source] + top - _pages[source].Margins.Top;
            double virtualBottom = _starts[source] + bottom - _pages[source].Margins.Top;
            int target = FindWindow(virtualTop);
            for (; target < _pages.Count && _starts[target] < virtualBottom - Epsilon; target++) {
                if (_starts[target + 1] <= virtualTop + Epsilon) continue;
                if ((visual is HtmlRenderNamedDestination || visual is HtmlRenderBookmarkAnchor)
                    && !_ownedNavigation.Add(visual)) break;
                double dy = OffsetY(source, target);
                double dx = _pages[target].Margins.Left - _pages[source].Margins.Left;
                Count();
                HtmlRenderVisual projected = visual.TranslatePaint(dx, dy, visual.PaintOrder);
                if (visual is HtmlRenderText && !_ownedText.Add(visual)) {
                    Count();
                    projected = new HtmlRenderLogicalTextGroup(string.Empty,
                        projected.X, projected.Y, projected.Width, projected.Height,
                        new[] { projected }, projected.PaintOrder, projected.Source,
                        projected.LayoutY, projected.LayoutHeight);
                }
                double visibleTop = Math.Max(top + dy, _pages[target].Margins.Top);
                double visibleBottom = Math.Min(bottom + dy, _pages[target].Height - _pages[target].Margins.Bottom);
                if (visibleBottom <= visibleTop + Epsilon) continue;
                if (visibleTop > paintBounds.Y + dy + Epsilon || visibleBottom < paintBounds.Bottom + dy - Epsilon) {
                    Count();
                    projected = ClipToPhysicalWindow(projected, paintBounds.X + dx, paintBounds.Right + dx,
                        visibleTop, visibleBottom, transform, dx, dy);
                }
                routed[target] = projected;
            }
            return routed;
        }

        private static bool CanMove(HtmlRenderVisual visual) =>
            Math.Abs(visual.RelativePaintOffsetY) > Epsilon
            && visual is not HtmlRenderLayoutBox
            && visual is not HtmlRenderLogicalTextGroup { IsFlowAnchor: true };

        private static void Constrain(HtmlRenderVisual visual, OfficeTransform transform, ref double top, ref double bottom) {
            HtmlRenderRectangle? bounds = null;
            if (visual is HtmlRenderClipGroup { ClipVertical: true } clip) {
                double y = clip.ClipY + (clip.IsFlowFragment ? clip.RelativePaintOffsetY : 0D);
                if (clip.ClipHorizontal || Math.Abs(transform.M12) < Epsilon) {
                    bounds = HtmlRenderRectangle.Transform(transform, new HtmlRenderRectangle(
                        clip.ClipX, y, clip.ClipWidth, clip.ClipHeight));
                }
                // A rotated one-axis clip has unbounded support on the other
                // axis. Retain its exact scene clip rather than guessing a box.
            } else if (visual is HtmlRenderPathClipGroup path) {
                bounds = PaintBounds(path, transform);
            }
            if (bounds.HasValue) {
                top = Math.Max(top, bounds.Value.Y);
                bottom = Math.Min(bottom, bounds.Value.Bottom);
            }
        }

        private static HtmlRenderRectangle PaintBounds(HtmlRenderVisual visual, OfficeTransform transform) {
            double overhang = visual is HtmlRenderText text ? text.PaintTopOverflow : 0D;
            double width = visual is HtmlRenderText measured ? Math.Max(measured.Width, measured.TextPaintWidth ?? measured.Width) : visual.Width;
            return HtmlRenderRectangle.Transform(transform,
                new HtmlRenderRectangle(visual.X, visual.Y - overhang, width, visual.Height + overhang));
        }

        private static HtmlRenderVisual ClipToPhysicalWindow(HtmlRenderVisual visual, double left, double right,
            double top, double bottom, OfficeTransform transform, double dx, double dy) {
            if (transform == OfficeTransform.Identity) {
                return new HtmlRenderClipGroup(left, top, Math.Max(0.01D, right - left), bottom - top,
                    false, true, new[] { visual }, visual.PaintOrder, visual.Source);
            }
            // The wrapper remains inside its authored effect groups. Express
            // the physical page window in that coordinate space, then let the
            // existing effects transform both paint and clipping together.
            OfficeTransform translatedTransform = OfficeTransform.Translate(-dx, -dy).Then(transform)
                .Then(OfficeTransform.Translate(dx, dy));
            OfficeTransform inverse = translatedTransform.Invert();
            OfficePoint[] points = {
                inverse.TransformPoint(new OfficePoint(left, top)),
                inverse.TransformPoint(new OfficePoint(right, top)),
                inverse.TransformPoint(new OfficePoint(right, bottom)),
                inverse.TransformPoint(new OfficePoint(left, bottom))
            };
            OfficeClipPath path = OfficeClipPath.Path(OfficePathCommand.MoveTo(points[0].X, points[0].Y),
                OfficePathCommand.LineTo(points[1].X, points[1].Y), OfficePathCommand.LineTo(points[2].X, points[2].Y),
                OfficePathCommand.LineTo(points[3].X, points[3].Y), OfficePathCommand.Close());
            return new HtmlRenderPathClipGroup(points.Min(point => point.X), points.Min(point => point.Y),
                path, new[] { visual }, visual.PaintOrder, visual.Source);
        }

        private double OffsetY(int source, int target) =>
            _starts[source] - _starts[target] + _pages[target].Margins.Top - _pages[source].Margins.Top;

        private int FindWindow(double y) {
            int low = 0, high = _pages.Count - 1;
            while (low < high) {
                int middle = low + (high - low) / 2;
                if (_starts[middle + 1] <= y + Epsilon) low = middle + 1;
                else high = middle;
            }
            return low;
        }

        private static Dictionary<int, HtmlRenderVisual> AtSource(HtmlRenderVisual visual, int source) =>
            new() { [source] = visual };

        private void Count() {
            _engine.CheckCancellation();
            if (++_materialized > _engine._options.MaxProjectedVisuals) {
                throw new InvalidOperationException("HTML relative paint projection exceeded MaxProjectedVisuals "
                    + _engine._options.MaxProjectedVisuals + ".");
            }
        }
    }
}
