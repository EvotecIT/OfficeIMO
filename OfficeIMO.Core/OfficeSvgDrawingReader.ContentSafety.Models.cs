using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.ContentSafety;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private sealed class SvgContentSafetyDocument {
        internal SvgContentSafetyDocument(
            XDocument document,
            XElement root,
            int maximumElements,
            double maximumViewportDimension,
            double maximumViewportPixels,
            double viewX,
            double viewY,
            double viewWidth,
            double viewHeight,
            double viewportWidth,
            double viewportHeight,
            int maximumVisualComparisons,
            long maximumVisualPixels) {
            Document = document;
            Root = root;
            MaximumElements = maximumElements;
            MaximumViewportDimension = maximumViewportDimension;
            MaximumViewportPixels = maximumViewportPixels;
            ViewX = viewX;
            ViewY = viewY;
            ViewWidth = viewWidth;
            ViewHeight = viewHeight;
            ViewportWidth = viewportWidth;
            ViewportHeight = viewportHeight;
            MaximumVisualComparisons = maximumVisualComparisons;
            MaximumVisualPixels = maximumVisualPixels;
        }

        internal XDocument Document { get; }
        internal XElement Root { get; }
        internal int MaximumElements { get; }
        internal double MaximumViewportDimension { get; }
        internal double MaximumViewportPixels { get; }
        internal double ViewX { get; }
        internal double ViewY { get; }
        internal double ViewWidth { get; }
        internal double ViewHeight { get; }
        internal double ViewportWidth { get; }
        internal double ViewportHeight { get; }
        internal int MaximumVisualComparisons { get; }
        internal long MaximumVisualPixels { get; }
    }

    private sealed class SvgContentSafetyCandidate {
        private bool _hasBounds;
        private double _left;
        private double _top;
        private double _right;
        private double _bottom;

        internal SvgContentSafetyCandidate(
            XText sourceText,
            XText computedText,
            XElement computedElement,
            int elementIndex,
            int textNodeIndex,
            string location,
            bool isNativeSvg) {
            SourceText = sourceText;
            ComputedText = computedText;
            ComputedElement = computedElement;
            ElementIndex = elementIndex;
            TextNodeIndex = textNodeIndex;
            Location = location;
            IsNativeSvg = isNativeSvg;
            Target = new SvgContentSafetyTarget(sourceText);
        }

        internal XText SourceText { get; }
        internal XText ComputedText { get; }
        internal XElement ComputedElement { get; }
        internal int ElementIndex { get; }
        internal int TextNodeIndex { get; }
        internal string Location { get; }
        internal bool IsNativeSvg { get; }
        internal SvgContentSafetyTarget Target { get; }
        internal SvgPaintContext Style { get; set; }
        internal bool HasStyle { get; set; }
        internal bool HasBounds => _hasBounds;
        internal double Left => _left;
        internal double Top => _top;
        internal double Right => _right;
        internal double Bottom => _bottom;
        internal double MaximumEffectiveFontSize { get; private set; }

        internal void Include(SvgTextRun run) {
            double y = run.Baseline - run.FontSize;
            double height = run.FontSize * 1.25D;
            double strokeExtent = 0D;
            if (HasStyle && Style.StrokeWidth > 0D &&
                (Style.Stroke.HasValue || Style.StrokeGradient != null || Style.StrokeRadialGradient != null ||
                 Style.StrokeDeferredGradient != null || Style.StrokePattern != null)) {
                strokeExtent = Style.StrokeWidth / 2D;
                if (Style.LineJoin == OfficeStrokeLineJoin.Miter) {
                    strokeExtent *= Math.Max(1D, Style.MiterLimit);
                }
            }
            OfficeTransform transform = Math.Abs(run.RotationDegrees) <= 0.0000001D
                ? run.Transform
                : OfficeTransform.RotateDegrees(
                    run.RotationDegrees,
                    run.RotationCenterX,
                    run.RotationCenterY).Then(run.Transform);
            (double Left, double Top, double Right, double Bottom) bounds =
                transform.TransformRectangleBounds(
                    run.X - strokeExtent,
                    y - strokeExtent,
                    run.Width + strokeExtent * 2D,
                    height + strokeExtent * 2D);
            if (strokeExtent > 0D) {
                bounds = (
                    bounds.Left - strokeExtent,
                    bounds.Top - strokeExtent,
                    bounds.Right + strokeExtent,
                    bounds.Bottom + strokeExtent);
            }
            double horizontalScale = Math.Sqrt(transform.M11 * transform.M11 + transform.M12 * transform.M12);
            double verticalScale = Math.Sqrt(transform.M21 * transform.M21 + transform.M22 * transform.M22);
            double minimumScale = Math.Min(horizontalScale, verticalScale);
            MaximumEffectiveFontSize = Math.Max(MaximumEffectiveFontSize, run.FontSize * minimumScale);
            if (!_hasBounds) {
                _left = bounds.Left;
                _top = bounds.Top;
                _right = bounds.Right;
                _bottom = bounds.Bottom;
                _hasBounds = true;
                return;
            }
            _left = Math.Min(_left, bounds.Left);
            _top = Math.Min(_top, bounds.Top);
            _right = Math.Max(_right, bounds.Right);
            _bottom = Math.Max(_bottom, bounds.Bottom);
        }
    }

    private sealed class SvgContentSafetyTextObserver {
        private readonly IReadOnlyDictionary<XText, SvgContentSafetyCandidate> _byText;
        private readonly Dictionary<SvgTextRun, SvgContentSafetyCandidate> _byRun =
            new Dictionary<SvgTextRun, SvgContentSafetyCandidate>();

        internal SvgContentSafetyTextObserver(IReadOnlyDictionary<XText, SvgContentSafetyCandidate> byText) {
            _byText = byText;
        }

        internal void Associate(
            XText text,
            SvgPaintContext style,
            IList<SvgTextRun> runs,
            int firstRun) {
            if (!_byText.TryGetValue(text, out SvgContentSafetyCandidate? candidate)) return;
            candidate.Style = style;
            candidate.HasStyle = true;
            for (int index = firstRun; index < runs.Count; index++) _byRun[runs[index]] = candidate;
        }

        internal void AssociateSubtree(XElement element, SvgPaintContext style) {
            foreach (XText text in element.DescendantNodes().OfType<XText>()) {
                if (!_byText.TryGetValue(text, out SvgContentSafetyCandidate? candidate)) continue;
                candidate.Style = style;
                candidate.HasStyle = true;
            }
        }

        internal void FinalizeRuns(IEnumerable<SvgTextRun> runs) {
            foreach (SvgTextRun run in runs) {
                if (_byRun.TryGetValue(run, out SvgContentSafetyCandidate? candidate)) candidate.Include(run);
            }
        }
    }

    private sealed class SvgContentSafetyTarget {
        private readonly XText _text;

        internal SvgContentSafetyTarget(XText text) {
            _text = text;
        }

        internal void Remove(IEnumerable<OfficeContentSafetyFinding> findings) {
            OfficeContentSafetyFinding[] selected = findings.ToArray();
            if (selected.Any(item => item.Kind != OfficeContentConcealmentKind.NonPrintingUnicode)) {
                _text.Remove();
                return;
            }

            string value = _text.Value;
            foreach (OfficeContentSafetyFinding finding in selected.OrderByDescending(item => item.SourceTextOffset ?? -1)) {
                if (!finding.SourceTextOffset.HasValue || !finding.SourceTextLength.HasValue) {
                    throw new InvalidOperationException("The selected SVG Unicode finding has no exact source range.");
                }
                int offset = finding.SourceTextOffset.Value;
                int length = finding.SourceTextLength.Value;
                if (offset < 0 || length <= 0 || offset > value.Length - length) {
                    throw new InvalidOperationException("The selected SVG Unicode finding no longer matches its source text.");
                }
                value = value.Remove(offset, length);
            }
            _text.Value = value;
        }
    }

    private readonly struct SvgContentSafetyConcealment {
        internal SvgContentSafetyConcealment(
            OfficeContentConcealmentKind kind,
            string evidence,
            OfficeContentSafetyRisk risk = OfficeContentSafetyRisk.ContextDependent) {
            Kind = kind;
            Evidence = evidence;
            Risk = risk;
        }

        internal OfficeContentConcealmentKind Kind { get; }
        internal string Evidence { get; }
        internal OfficeContentSafetyRisk Risk { get; }
    }

    private readonly struct SvgVisualComparison {
        internal SvgVisualComparison(int changedPixels, double maximumContrastRatio, bool hasTransparentBackdrop) {
            ChangedPixels = changedPixels;
            MaximumContrastRatio = maximumContrastRatio;
            HasTransparentBackdrop = hasTransparentBackdrop;
        }

        internal int ChangedPixels { get; }
        internal double MaximumContrastRatio { get; }
        internal bool HasTransparentBackdrop { get; }
    }

    private static bool IsNativeSvgElement(XElement element, XNamespace svgNamespace) =>
        element.Name.Namespace == svgNamespace;

    private static bool IsSupportedSvgElementNamespace(XElement element) =>
        element.Name.NamespaceName.Length == 0 ||
        element.Name.NamespaceName.Equals("http://www.w3.org/2000/svg", StringComparison.Ordinal);
}
