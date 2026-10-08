using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Describes the geometric precision used for an HTML scene hit.</summary>
public enum HtmlRenderHitTestAccuracy {
    /// <summary>Axis-aligned visual bounds were tested after applying retained transforms.</summary>
    TransformedBounds,
    /// <summary>Visual bounds and every retained rectangular, rounded, or path clip were tested.</summary>
    ClipAwareTransformedBounds
}

/// <summary>Limits and filters for one retained-scene hit test.</summary>
public sealed class HtmlRenderHitTestOptions {
    /// <summary>Default maximum matches returned by one query.</summary>
    public const int DefaultMaximumResults = 32;
    /// <summary>Default maximum scene nodes visited by one query.</summary>
    public const int DefaultMaximumVisitedVisuals = 100_000;
    /// <summary>Default maximum flattened clip-path line segments evaluated by one query.</summary>
    public const int DefaultMaximumPathSegments = 100_000;

    /// <summary>Maximum topmost-first matches returned.</summary>
    public int MaximumResults { get; set; } = DefaultMaximumResults;
    /// <summary>Maximum scene nodes inspected before the report is marked truncated.</summary>
    public int MaximumVisitedVisuals { get; set; } = DefaultMaximumVisitedVisuals;
    /// <summary>Maximum flattened freeform clip-path line segments evaluated across one query.</summary>
    public int MaximumPathSegments { get; set; } = DefaultMaximumPathSegments;
    /// <summary>Return only links and retained form controls.</summary>
    public bool InteractiveOnly { get; set; }
    /// <summary>Include grouping nodes in addition to leaf visuals and form controls.</summary>
    public bool IncludeContainers { get; set; }

    /// <summary>Returns an independent options copy with another result limit.</summary>
    public HtmlRenderHitTestOptions WithMaximumResults(int maximumResults) {
        var copy = Clone();
        copy.MaximumResults = maximumResults;
        copy.Validate();
        return copy;
    }

    internal HtmlRenderHitTestOptions Clone() => new HtmlRenderHitTestOptions {
        MaximumResults = MaximumResults,
        MaximumVisitedVisuals = MaximumVisitedVisuals,
        MaximumPathSegments = MaximumPathSegments,
        InteractiveOnly = InteractiveOnly,
        IncludeContainers = IncludeContainers
    };

    internal void Validate() {
        if (MaximumResults < 1 || MaximumResults > 10_000) {
            throw new ArgumentOutOfRangeException(nameof(MaximumResults), "MaximumResults must be between 1 and 10000.");
        }
        if (MaximumVisitedVisuals < 1 || MaximumVisitedVisuals > 10_000_000) {
            throw new ArgumentOutOfRangeException(nameof(MaximumVisitedVisuals), "MaximumVisitedVisuals must be between 1 and 10000000.");
        }
        if (MaximumPathSegments < 1 || MaximumPathSegments > 10_000_000) {
            throw new ArgumentOutOfRangeException(nameof(MaximumPathSegments), "MaximumPathSegments must be between 1 and 10000000.");
        }
    }
}

/// <summary>One visual matched by a retained-scene hit test.</summary>
public sealed class HtmlRenderHitTestResult {
    internal HtmlRenderHitTestResult(HtmlRenderSurface surface, HtmlRenderVisual visual,
        HtmlRenderRectangle bounds, HtmlRenderPoint point, HtmlRenderSourcePoint? sourcePoint,
        int depth, HtmlRenderHitTestAccuracy accuracy) {
        Surface = surface;
        Visual = visual;
        Bounds = bounds;
        Point = point;
        SourcePoint = sourcePoint;
        Depth = depth;
        Accuracy = accuracy;
    }

    /// <summary>Surface containing the match.</summary>
    public HtmlRenderSurface Surface { get; }
    /// <summary>Matched retained visual.</summary>
    public HtmlRenderVisual Visual { get; }
    /// <summary>Axis-aligned output bounds after retained transforms.</summary>
    public HtmlRenderRectangle Bounds { get; }
    /// <summary>Queried output point.</summary>
    public HtmlRenderPoint Point { get; }
    /// <summary>Mapped source point, or null when the point falls outside every source placement.</summary>
    public HtmlRenderSourcePoint? SourcePoint { get; }
    /// <summary>Zero-based depth in the retained scene.</summary>
    public int Depth { get; }
    /// <summary>Geometric precision used for the match.</summary>
    public HtmlRenderHitTestAccuracy Accuracy { get; }
    /// <summary>Whether the match owns a hyperlink or standard form control.</summary>
    public bool IsInteractive => Visual.LinkUri != null || Visual is HtmlRenderFormField;
}

/// <summary>Bounded topmost-first hit-test result.</summary>
public sealed class HtmlRenderHitTestReport {
    private readonly ReadOnlyCollection<HtmlRenderHitTestResult> _matches;

    internal HtmlRenderHitTestReport(IEnumerable<HtmlRenderHitTestResult> matches,
        int visitedVisualCount, bool truncated) {
        _matches = new List<HtmlRenderHitTestResult>(matches).AsReadOnly();
        VisitedVisualCount = visitedVisualCount;
        IsTruncated = truncated;
    }

    /// <summary>Topmost-first matching visuals.</summary>
    public IReadOnlyList<HtmlRenderHitTestResult> Matches => _matches;
    /// <summary>Number of retained scene nodes inspected.</summary>
    public int VisitedVisualCount { get; }
    /// <summary>Whether a result, visited-node, or clip-path segment limit stopped traversal.</summary>
    public bool IsTruncated { get; }
}

internal static class HtmlRenderHitTester {
    internal static HtmlRenderHitTestReport HitTest(HtmlRenderSurface surface, HtmlRenderPoint point,
        HtmlRenderHitTestOptions? options, CancellationToken cancellationToken) {
        if (surface == null) throw new ArgumentNullException(nameof(surface));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderHitTestOptions resolved = options?.Clone() ?? new HtmlRenderHitTestOptions();
        resolved.Validate();
        if (!surface.Bounds.Contains(point)) {
            return new HtmlRenderHitTestReport(Array.Empty<HtmlRenderHitTestResult>(), 0, false);
        }

        var state = new State(surface, point, resolved, cancellationToken);
        Traverse(state, surface.Page.Scene, OfficeTransform.Identity, 0, clipsApplied: false);
        return new HtmlRenderHitTestReport(state.Matches, state.Visited, state.Truncated);
    }

    private static void Traverse(State state, IReadOnlyList<HtmlRenderVisual> visuals,
        OfficeTransform transform, int depth, bool clipsApplied) {
        for (int index = visuals.Count - 1; index >= 0 && !state.Stop; index--) {
            state.CancellationToken.ThrowIfCancellationRequested();
            HtmlRenderVisual visual = visuals[index];
            if (++state.Visited > state.Options.MaximumVisitedVisuals) {
                state.Truncated = true;
                return;
            }

            OfficeTransform visualTransform = visual is HtmlRenderEffectGroup effect
                ? effect.Transform.Then(transform)
                : transform;
            if (!TryMapPoint(state.Point, visualTransform, out HtmlRenderPoint localPoint)) continue;

            bool visualClipped = clipsApplied;
            IReadOnlyList<HtmlRenderVisual>? children = GetChildren(visual);
            if (visual is HtmlRenderClipGroup rectangleClip) {
                if ((rectangleClip.ClipHorizontal &&
                     (localPoint.X < rectangleClip.ClipX || localPoint.X > rectangleClip.ClipX + rectangleClip.ClipWidth)) ||
                    (rectangleClip.ClipVertical &&
                     (localPoint.Y < rectangleClip.ClipY || localPoint.Y > rectangleClip.ClipY + rectangleClip.ClipHeight))) {
                    continue;
                }
                visualClipped = true;
            } else if (visual is HtmlRenderPathClipGroup pathClip) {
                if (!Contains(pathClip.ClipPath,
                        localPoint.X - pathClip.ClipX,
                        localPoint.Y - pathClip.ClipY,
                        state)) {
                    continue;
                }
                visualClipped = true;
            } else if (visual is HtmlRenderEffectGroup invisible && invisible.Opacity <= 0D) {
                continue;
            }

            if (children != null) {
                Traverse(state, children, visualTransform, depth + 1, visualClipped);
            }
            if (state.Stop) return;

            bool isContainer = children != null;
            bool interactive = visual.LinkUri != null || visual is HtmlRenderFormField;
            if ((isContainer && !state.Options.IncludeContainers && !interactive) ||
                (state.Options.InteractiveOnly && !interactive)) {
                continue;
            }

            HtmlRenderRectangle localBounds = new HtmlRenderRectangle(visual.X, visual.Y, visual.Width, visual.Height);
            if (!localBounds.Contains(localPoint)) continue;
            HtmlRenderRectangle outputBounds = HtmlRenderRectangle.Transform(visualTransform, localBounds);
            state.Surface.TryMapToSource(state.Point, out HtmlRenderSourcePoint? sourcePoint);
            state.Matches.Add(new HtmlRenderHitTestResult(
                state.Surface,
                visual,
                outputBounds,
                state.Point,
                sourcePoint,
                depth,
                visualClipped
                    ? HtmlRenderHitTestAccuracy.ClipAwareTransformedBounds
                    : HtmlRenderHitTestAccuracy.TransformedBounds));
            if (state.Matches.Count >= state.Options.MaximumResults) {
                state.Truncated = HasPotentialAdditionalMatches(visuals, index) || depth > 0;
                return;
            }
        }
    }

    private static bool HasPotentialAdditionalMatches(IReadOnlyList<HtmlRenderVisual> visuals, int currentIndex) =>
        currentIndex > 0;

    private static IReadOnlyList<HtmlRenderVisual>? GetChildren(HtmlRenderVisual visual) => visual switch {
        HtmlRenderClipGroup group => group.Visuals,
        HtmlRenderPathClipGroup group => group.Visuals,
        HtmlRenderEffectGroup group => group.Visuals,
        HtmlRenderSemanticGroup group => group.Visuals,
        HtmlRenderLayoutRegion group => group.Visuals,
        HtmlRenderLogicalTextGroup group => group.Visuals,
        HtmlRenderFormField group => group.Visuals,
        _ => null
    };

    private static bool TryMapPoint(HtmlRenderPoint outputPoint, OfficeTransform transform,
        out HtmlRenderPoint localPoint) {
        if (!transform.TryInvert(out OfficeTransform inverse)) {
            localPoint = default;
            return false;
        }
        OfficePoint point = inverse.TransformPoint(outputPoint.ToOfficePoint());
        localPoint = new HtmlRenderPoint(point.X, point.Y);
        return true;
    }

    private static bool Contains(OfficeClipPath path, double x, double y, State state) {
        if (x < 0D || y < 0D || x > path.Width || y > path.Height) return false;
        switch (path.Kind) {
            case OfficeClipPathKind.Empty:
                return false;
            case OfficeClipPathKind.Rectangle:
                return true;
            case OfficeClipPathKind.RoundedRectangle:
                return ContainsRoundedRectangle(path.Width, path.Height, path.CornerRadius, x, y);
            case OfficeClipPathKind.Path:
                return ContainsPath(path, x, y, state);
            default:
                return false;
        }
    }

    private static bool ContainsRoundedRectangle(double width, double height, double radius, double x, double y) {
        if (radius <= 0D || (x >= radius && x <= width - radius) || (y >= radius && y <= height - radius)) {
            return true;
        }
        double centerX = x < radius ? radius : width - radius;
        double centerY = y < radius ? radius : height - radius;
        double dx = x - centerX;
        double dy = y - centerY;
        return dx * dx + dy * dy <= radius * radius;
    }

    private static bool ContainsPath(OfficeClipPath path, double x, double y, State state) {
        IReadOnlyList<IReadOnlyList<OfficePoint>> contours = Flatten(path.Commands, state);
        if (state.Truncated) return false;
        if (path.FillRule == OfficeFillRule.EvenOdd) {
            bool inside = false;
            foreach (IReadOnlyList<OfficePoint> contour in contours) {
                for (int current = 0, previous = contour.Count - 1; current < contour.Count; previous = current++) {
                    OfficePoint a = contour[current];
                    OfficePoint b = contour[previous];
                    if (((a.Y > y) != (b.Y > y)) &&
                        x < (b.X - a.X) * (y - a.Y) / (b.Y - a.Y) + a.X) {
                        inside = !inside;
                    }
                }
            }
            return inside;
        }

        int winding = 0;
        foreach (IReadOnlyList<OfficePoint> contour in contours) {
            for (int current = 0, previous = contour.Count - 1; current < contour.Count; previous = current++) {
                OfficePoint a = contour[previous];
                OfficePoint b = contour[current];
                if (a.Y <= y) {
                    if (b.Y > y && IsLeft(a, b, x, y) > 0D) winding++;
                } else if (b.Y <= y && IsLeft(a, b, x, y) < 0D) {
                    winding--;
                }
            }
        }
        return winding != 0;
    }

    private static IReadOnlyList<IReadOnlyList<OfficePoint>> Flatten(IReadOnlyList<OfficePathCommand> commands, State state) {
        const int curveSegments = 16;
        var contours = new List<IReadOnlyList<OfficePoint>>();
        List<OfficePoint>? contour = null;
        OfficePoint current = default;
        OfficePoint start = default;
        foreach (OfficePathCommand command in commands) {
            state.CancellationToken.ThrowIfCancellationRequested();
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    if (contour != null && contour.Count > 1) contours.Add(contour.AsReadOnly());
                    contour = new List<OfficePoint> { command.Point };
                    current = start = command.Point;
                    break;
                case OfficePathCommandKind.LineTo:
                    if (!AddPathSegments(state, 1)) return contours.AsReadOnly();
                    contour!.Add(command.Point);
                    current = command.Point;
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    if (!AddPathSegments(state, curveSegments)) return contours.AsReadOnly();
                    for (int step = 1; step <= curveSegments; step++) {
                        double t = step / (double)curveSegments;
                        double inverse = 1D - t;
                        contour!.Add(new OfficePoint(
                            inverse * inverse * current.X + 2D * inverse * t * command.ControlPoint1.X + t * t * command.Point.X,
                            inverse * inverse * current.Y + 2D * inverse * t * command.ControlPoint1.Y + t * t * command.Point.Y));
                    }
                    current = command.Point;
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    if (!AddPathSegments(state, curveSegments)) return contours.AsReadOnly();
                    for (int step = 1; step <= curveSegments; step++) {
                        double t = step / (double)curveSegments;
                        double inverse = 1D - t;
                        contour!.Add(new OfficePoint(
                            inverse * inverse * inverse * current.X + 3D * inverse * inverse * t * command.ControlPoint1.X +
                            3D * inverse * t * t * command.ControlPoint2.X + t * t * t * command.Point.X,
                            inverse * inverse * inverse * current.Y + 3D * inverse * inverse * t * command.ControlPoint1.Y +
                            3D * inverse * t * t * command.ControlPoint2.Y + t * t * t * command.Point.Y));
                    }
                    current = command.Point;
                    break;
                case OfficePathCommandKind.Close:
                    if (contour != null && contour.Count > 1 && contour[contour.Count - 1] != start) {
                        if (!AddPathSegments(state, 1)) return contours.AsReadOnly();
                        contour.Add(start);
                    }
                    current = start;
                    break;
            }
        }
        if (contour != null && contour.Count > 1) contours.Add(contour.AsReadOnly());
        return contours.AsReadOnly();
    }

    private static bool AddPathSegments(State state, int count) {
        if (state.PathSegments > state.Options.MaximumPathSegments - count) {
            state.Truncated = true;
            return false;
        }
        state.PathSegments += count;
        return true;
    }

    private static double IsLeft(OfficePoint a, OfficePoint b, double x, double y) =>
        (b.X - a.X) * (y - a.Y) - (x - a.X) * (b.Y - a.Y);

    private sealed class State {
        internal State(HtmlRenderSurface surface, HtmlRenderPoint point,
            HtmlRenderHitTestOptions options, CancellationToken cancellationToken) {
            Surface = surface;
            Point = point;
            Options = options;
            CancellationToken = cancellationToken;
        }

        internal HtmlRenderSurface Surface { get; }
        internal HtmlRenderPoint Point { get; }
        internal HtmlRenderHitTestOptions Options { get; }
        internal CancellationToken CancellationToken { get; }
        internal List<HtmlRenderHitTestResult> Matches { get; } = new List<HtmlRenderHitTestResult>();
        internal int Visited { get; set; }
        internal int PathSegments { get; set; }
        internal bool Truncated { get; set; }
        internal bool Stop => Truncated || Matches.Count >= Options.MaximumResults;
    }
}
