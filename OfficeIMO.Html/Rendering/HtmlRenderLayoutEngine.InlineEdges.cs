using AngleSharp.Dom;

namespace OfficeIMO.Html;

// Each layout pass owns its scopes. Painted glyph/atomic widths remain positive;
// these signed advances describe the space consumed by their enclosing inline boxes.
internal sealed class HtmlInlineEdgeScope {
    internal HtmlInlineEdgeScope(IElement owner, HtmlRenderBoxStyle style) {
        Owner = owner;
        Style = style;
        Start = style.MarginLeft + style.PaddingLeft + style.BorderLeftWidth;
        End = style.MarginRight + style.PaddingRight + style.BorderRightWidth;
        Clone = style.BoxDecorationBreak == "clone";
    }
    internal IElement Owner { get; }
    internal HtmlRenderBoxStyle Style { get; }
    internal double Start { get; }
    internal double End { get; }
    internal bool Clone { get; }
    internal object? FirstContent { get; set; }
    internal object? LastContent { get; set; }
    internal bool Closed { get; set; }
}

internal sealed class HtmlInlineEdgeBoundary {
    internal HtmlInlineEdgeBoundary(HtmlInlineEdgeScope scope, bool opening) {
        Scope = scope;
        Opening = opening;
    }
    internal HtmlInlineEdgeScope Scope { get; }
    internal bool Opening { get; }
}

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly HashSet<IElement> _reportedInlineEdgeFallbacks = new();
    private readonly Dictionary<HtmlInlineEdgeScope, (double Left, double Right)> _currentInlineEdgeGeometry = new();
    private HtmlInlineEdgeScope? ResolveInlineEdgeScope(IElement element, HtmlRenderBoxStyle style, HtmlRenderBoxStyle parentStyle) {
        if (style.Display != "inline" || style.MarginLeft == 0D && style.MarginRight == 0D && style.HorizontalInsets == 0D) return null;
        if (style.Direction == "rtl" || parentStyle.Direction == "rtl" || style.WritingMode != "horizontal-tb") {
            if (_reportedInlineEdgeFallbacks.Add(element)) _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.InlinePaintEffectUnsupported,
                "Horizontal inline box advances require a left-to-right horizontal formatting context; decoration paint retains its diagnosed fallback.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                "inline-edges;direction=" + parentStyle.Direction + ";writing-mode=" + style.WritingMode,
                OfficeConversionLossKind.Approximation);
            return null;
        }
        return new HtmlInlineEdgeScope(element, style);
    }

    private static HtmlInlineRun CreateInlineEdgeBoundary(HtmlInlineEdgeScope scope, bool opening,
        string? link, double paintOffsetX, double paintOffsetY) =>
        new(string.Empty, scope.Style, link, HtmlRenderStyleResolver.DescribeSource(scope.Owner),
            paintOffsetX, paintOffsetY, scope.Owner) {
            InlineEdgeBoundary = new HtmlInlineEdgeBoundary(scope, opening)
        };

    private static void PrepareInlineEdgeScopes(IReadOnlyList<HtmlInlineRun> runs) {
        var scopes = new List<HtmlInlineEdgeScope>();
        foreach (HtmlInlineRun run in runs) {
            HtmlInlineEdgeBoundary? boundary = run.InlineEdgeBoundary;
            if (boundary?.Opening == true) {
                boundary.Scope.FirstContent = null;
                boundary.Scope.LastContent = null;
                boundary.Scope.Closed = false;
                scopes.Add(boundary.Scope);
            }
            run.InlineEdgeScopes = scopes.ToArray();
            run.InlineTokenEndsRun = true;
            if (boundary?.Opening == false && scopes.Count > 0) scopes.RemoveAt(scopes.Count - 1);
        }
        double followingClosingAdvance = 0D;
        for (int index = runs.Count - 1; index >= 0; index--) {
            HtmlInlineRun candidate = runs[index];
            candidate.InlineClosingAdvance = followingClosingAdvance;
            if (candidate.InlineEdgeBoundary is { Opening: false } closing) {
                if (!closing.Scope.Clone) followingClosingAdvance += closing.Scope.End;
            } else if (candidate.RunningStringElement == null && candidate.RunningElementAssignment == null
                && !candidate.IsFlowMarker && candidate.PositionedMarkerElement == null) {
                followingClosingAdvance = 0D;
            }
        }
    }

    private static void MarkSkippedInlineEdges(HtmlInlineRun run) {
        foreach (HtmlInlineEdgeScope scope in run.InlineEdgeScopes) {
            scope.FirstContent ??= scope; // A sliced continuation has already consumed its real start.
        }
    }

    private static bool ProcessInlineEdgeBoundary(HtmlInlineRun run, InlineLine line) {
        if (run.InlineEdgeBoundary == null) return false;
        if (!run.InlineEdgeBoundary.Opening) line.CloseInlineScope(run);
        return true;
    }

    private sealed partial class InlineLine {
        private readonly Dictionary<HtmlInlineEdgeScope, InlineEdgeFragment> _inlineEdges = new();

        internal void RecordInlineScopePositions(double start, IDictionary<HtmlInlineEdgeScope, (double Left, double Right)> positions) {
            positions.Clear();
            var slots = new Dictionary<InlineSegment, (double Before, double After)>();
            double cursor = start;
            foreach (InlineSegment segment in Segments) {
                slots[segment] = (cursor, cursor + segment.Advance);
                cursor += segment.Advance;
            }
            foreach (KeyValuePair<HtmlInlineEdgeScope, InlineEdgeFragment> entry in _inlineEdges) {
                HtmlInlineEdgeScope scope = entry.Key;
                InlineEdgeFragment fragment = entry.Value;
                double left = slots[fragment.First].Before;
                foreach (HtmlInlineEdgeScope ancestor in fragment.First.Run.InlineEdgeScopes) {
                    if (_inlineEdges.TryGetValue(ancestor, out InlineEdgeFragment? parent)
                        && ReferenceEquals(parent.First, fragment.First)
                        && (ancestor.Clone || ReferenceEquals(ancestor.FirstContent, fragment.First))) left += ancestor.Start;
                    if (ReferenceEquals(ancestor, scope)) break;
                }
                double right = slots[fragment.Last].After;
                foreach (HtmlInlineEdgeScope ancestor in fragment.Last.Run.InlineEdgeScopes) {
                    if (_inlineEdges.TryGetValue(ancestor, out InlineEdgeFragment? parent)
                        && ReferenceEquals(parent.Last, fragment.Last) && (ancestor.Clone || parent.Closed)) right -= ancestor.End;
                    if (ReferenceEquals(ancestor, scope)) break;
                }
                positions[scope] = (left, right);
            }
        }

        internal double PreviewAdvance(HtmlInlineRun run, double contentWidth, bool completesToken = true) {
            double result = Width + contentWidth;
            if (run.IsFlowMarker || run.RunningStringElement != null || run.RunningElementAssignment != null
                || run.PositionedMarkerElement != null) return result;
            foreach (HtmlInlineEdgeScope scope in run.InlineEdgeScopes) {
                if (_inlineEdges.ContainsKey(scope)) continue;
                if (scope.Clone || scope.FirstContent == null) result += scope.Start;
                if (scope.Clone) result += scope.End;
            }
            if (completesToken && run.InlineTokenEndsRun) result += run.InlineClosingAdvance;
            return result;
        }

        private void AddInlineEdges(InlineSegment segment) {
            if (segment.Run.RunningStringElement != null || segment.Run.RunningElementAssignment != null
                || segment.Run.PositionedMarkerElement != null || segment.Run.IsFlowMarker) return;
            foreach (HtmlInlineEdgeScope scope in segment.Run.InlineEdgeScopes) {
                scope.FirstContent ??= segment;
                if (!scope.Closed) scope.LastContent = segment;
                if (!_inlineEdges.TryGetValue(scope, out InlineEdgeFragment? fragment)) {
                    fragment = new InlineEdgeFragment(segment);
                    _inlineEdges.Add(scope, fragment);
                    if (scope.Clone || ReferenceEquals(scope.FirstContent, segment)) {
                        segment.LeadingAdvance += scope.Start;
                        Width += scope.Start;
                    }
                    if (scope.Clone) {
                        segment.TrailingAdvance += scope.End;
                        Width += scope.End;
                    }
                } else {
                    if (scope.Clone || fragment.Closed) {
                        fragment.Last.TrailingAdvance -= scope.End;
                        segment.TrailingAdvance += scope.End;
                    }
                    fragment.Last = segment;
                }
                if (scope.Closed && ReferenceEquals(scope.LastContent, segment) && !fragment.Closed) {
                    fragment.Closed = true;
                    if (!scope.Clone) {
                        segment.TrailingAdvance += scope.End;
                        Width += scope.End;
                    }
                }
            }
        }

        internal void CloseInlineScope(HtmlInlineRun boundaryRun) {
            HtmlInlineEdgeScope scope = boundaryRun.InlineEdgeBoundary!.Scope;
            if (!_inlineEdges.TryGetValue(scope, out InlineEdgeFragment? fragment)) {
                // An empty, decorated inline still creates a box and consumes its edges.
                if (ReferenceEquals(scope.FirstContent, scope)) return;
                var empty = new HtmlInlineRun(string.Empty, boundaryRun.Style, boundaryRun.LinkUri,
                    boundaryRun.Source, boundaryRun.PaintOffsetX, boundaryRun.PaintOffsetY, boundaryRun.OwnerElement) {
                    InlineEdgeScopes = boundaryRun.InlineEdgeScopes,
                    IsEmptyInlineBox = true
                };
                Add(new InlineSegment(string.Empty, 0D, empty));
                fragment = _inlineEdges[scope];
            }
            if (fragment.Closed) return;
            scope.Closed = true;
            scope.LastContent = fragment.Last;
            fragment.Closed = true;
            if (!scope.Clone) {
                fragment.Last.TrailingAdvance += scope.End;
                Width += scope.End;
            }
        }

        private void RebuildInlineEdges(InlineSegment removed) {
            foreach (HtmlInlineEdgeScope scope in removed.Run.InlineEdgeScopes) {
                if (ReferenceEquals(scope.FirstContent, removed)) {
                    scope.FirstContent = Segments.FirstOrDefault(segment => segment.Run.InlineEdgeScopes.Contains(scope));
                }
            }
            var closed = new HashSet<HtmlInlineEdgeScope>(_inlineEdges.Where(pair => pair.Value.Closed).Select(pair => pair.Key));
            _inlineEdges.Clear();
            Width = Segments.Sum(segment => segment.Width);
            foreach (InlineSegment segment in Segments) {
                segment.LeadingAdvance = 0D;
                segment.TrailingAdvance = 0D;
                AddInlineEdges(segment);
            }
            foreach (HtmlInlineEdgeScope scope in closed) {
                if (!_inlineEdges.TryGetValue(scope, out InlineEdgeFragment? fragment)) continue;
                fragment.Closed = true;
                if (!scope.Clone) {
                    fragment.Last.TrailingAdvance += scope.End;
                    Width += scope.End;
                }
            }
        }

        private sealed class InlineEdgeFragment {
            internal InlineEdgeFragment(InlineSegment first) { First = first; Last = first; }
            internal InlineSegment First { get; }
            internal InlineSegment Last { get; set; }
            internal bool Closed { get; set; }
        }
    }
}
