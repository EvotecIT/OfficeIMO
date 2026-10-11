using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    /// <summary>
    /// Collects a bounded set of completely visible ordinary stroke centres through the
    /// canonical content/form parser. Uncertain paint or clipping leaves text inference
    /// in charge; an intersection with the page alone cannot prove a complete ruling.
    /// </summary>
    internal IReadOnlyList<PdfTableRuling> GetTableRulings(PdfUnderstandingPageContext context) {
        context.ThrowIfCancellationRequested();
        if (context.ImagePlacements.Count > 0) return Array.Empty<PdfTableRuling>();
        (double width, double height) = GetVisualPageSize();
        var pageContentBudget = new PageContentBudget(this, context.CancellationToken);
        Matrix2D transform = GetVisualPageTransform();
        // This bounded profile declines pages with graphics effects. IsDefault alone
        // does not exclude unresolved soft masks or unmodelled graphics-state paint.
        foreach (PdfPageDrawingEffectTransition transition in GetGraphicsEffectTransitions(transform, height, pageContentBudget)) {
            context.ConsumeWork();
            PdfPageDrawingEffect effect = transition.Effect;
            if (effect.HasUnsupportedGraphicsState || effect.SoftMask != null ||
                effect.HasUnresolvedSoftMask || effect.BlendMode != OfficeBlendMode.Normal) return Array.Empty<PdfTableRuling>();
        }
        var rulings = new List<PdfTableRuling>();
        bool uncertain = false;
        (OfficeColor Color, double Width, double Opacity)? strokePaint = null;
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        string content = GetContentStreamContent(pageContentBudget);
        if (content.Length > 0) {
            CollectVisualPrimitivesAndForms(content, resources, transform, width, height, Visit,
                new HashSet<PdfStream>(), retainPrimitiveData: false,
                // Opaque paint can hide a ruling. Images/groups are outside this small
                // stroke-only profile rather than evidence for an authored grid.
                type3ImageVisitor: (_, _, _) => uncertain = true,
                type3GroupVisitor: (_, _, _, _, _) => uncertain = true,
                textOutputBudget: CreateTextOutputBudget(), pageContentBudget: pageContentBudget,
                contentOrderPrefix: PdfContentOrderKey.Root,
                scaleStrokeWidthWithTransform: true);
        }
        return uncertain ? Array.Empty<PdfTableRuling>() : rulings;

        void Visit(PdfPageVisualPrimitive primitive) {
            context.ConsumeWork();
            if (uncertain) return;
            if (primitive.ContentOrderKey?.Depth != 1) { uncertain = true; return; }
            if (primitive.HasFillPaint) { uncertain = true; return; }
            if (!primitive.HasStrokePaint) return;
            double opacity = primitive.StrokeOpacity ?? 1D;
            if (IsFinite(opacity) && opacity <= 0D) return;
            if (!primitive.StrokeColor.HasValue ||
                !IsFinite(primitive.StrokeWidth) || primitive.StrokeWidth <= 0D ||
                !IsFinite(opacity) ||
                primitive.StrokeGradient != null || primitive.StrokeRadialGradient != null || primitive.StrokeTilingPattern != null ||
                primitive.StrokeDashStyle != OfficeStrokeDashStyle.Solid || primitive.StrokeDashPattern is { Array.Count: > 0 }) {
                uncertain = true;
                return;
            }
            OfficeColor color = primitive.StrokeColor.Value;
            if (color.A == 0) return;
            // This root stroke-only profile has no painted background. White strokes
            // on the default white page do not establish visible cell boundaries.
            if (color.R == byte.MaxValue && color.G == byte.MaxValue && color.B == byte.MaxValue) { uncertain = true; return; }
            var paint = (color, primitive.StrokeWidth, opacity);
            // Geometry alone cannot establish that an earlier rule survived later
            // paint. Decline mixed stroke paint before clipping/deduplication: even
            // a partial, clipped or full-span white stroke can erase a black rule.
            if (strokePaint.HasValue && !strokePaint.Value.Equals(paint)) { uncertain = true; return; }
            strokePaint = paint;
            double pad = primitive.StrokeWidth / 2D;
            if (!Contained(primitive.X - pad, primitive.Y - pad, primitive.X + primitive.Width + pad,
                    primitive.Y + primitive.Height + pad, 0D, 0D, width, height)) return;
            if (primitive.ClipPath is PdfPageClipPath clip &&
                (!clip.IsExact || !clip.IsRectangle || clip.ContainsTextClipping ||
                 !Contained(primitive.X - pad, primitive.Y - pad, primitive.X + primitive.Width + pad,
                    primitive.Y + primitive.Height + pad, clip.X, clip.Y, clip.X + clip.Width, clip.Y + clip.Height))) return;

            if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) Add(primitive.X1, primitive.Y1, primitive.X2, primitive.Y2);
            else if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
                double right = primitive.X + primitive.Width, bottom = primitive.Y + primitive.Height;
                Add(primitive.X, primitive.Y, right, primitive.Y);
                Add(right, primitive.Y, right, bottom);
                Add(right, bottom, primitive.X, bottom);
                Add(primitive.X, bottom, primitive.X, primitive.Y);
            } else {
                // Paths may contain many independent straight subpaths. Never retain
                // the parser's borrowed command buffer beyond this visitor.
                OfficePoint? previous = null, start = null;
                foreach (OfficePathCommand command in primitive.PathCommands) {
                    context.ConsumeWork();
                    if (command.Kind == OfficePathCommandKind.MoveTo) previous = start = command.Point;
                    else if (command.Kind == OfficePathCommandKind.LineTo && previous.HasValue) {
                        Add(previous.Value.X, previous.Value.Y, command.Point.X, command.Point.Y);
                        previous = command.Point;
                    } else if (command.Kind == OfficePathCommandKind.Close && previous.HasValue && start.HasValue) {
                        Add(previous.Value.X, previous.Value.Y, start.Value.X, start.Value.Y);
                        previous = start;
                    } else { uncertain = true; return; }
                }
            }
        }

        void Add(double x1, double y1, double x2, double y2) {
            context.ConsumeWork();
            if (uncertain || !IsFinite(x1) || !IsFinite(y1) || !IsFinite(x2) || !IsFinite(y2)) return;
            (double X, double Y) first = ToReadingPoint(x1, y1), second = ToReadingPoint(x2, y2);
            if (!IsFinite(first.X) || !IsFinite(first.Y) || !IsFinite(second.X) || !IsFinite(second.Y)) return;
            double dx = Math.Abs(first.X - second.X), dy = Math.Abs(first.Y - second.Y);
            // Nonorthogonal decorative strokes cannot erase an earlier rule with the
            // same ordinary paint; differing paint has already declined this page.
            if ((dx > TableDetector.RulingTolerance && dy > TableDetector.RulingTolerance) ||
                (dx <= TableDetector.RulingTolerance && dy <= TableDetector.RulingTolerance)) return;
            if (rulings.Count >= TableDetector.MaximumRulingSegments) { uncertain = true; return; }
            rulings.Add(new PdfTableRuling(first.X, first.Y, second.X, second.Y));
        }

        (double X, double Y) ToReadingPoint(double x, double y) {
            (double X, double Y) user = PdfVisualCoordinateMapper.TransformVisualPointToUser(
                GetPageBoundaryBox(), GetRotationDegrees(), x, y, GetEffectiveUserUnit());
            return context.ReadingFrame is null ? user : context.ReadingFrame.ToFrame(user.X, user.Y);
        }
    }

    private static bool Contained(double left, double top, double right, double bottom,
        double outerLeft, double outerTop, double outerRight, double outerBottom) =>
        IsFinite(left) && IsFinite(top) && IsFinite(right) && IsFinite(bottom) &&
        IsFinite(outerLeft) && IsFinite(outerTop) && IsFinite(outerRight) && IsFinite(outerBottom) &&
        left >= outerLeft && top >= outerTop && right <= outerRight && bottom <= outerBottom;
}
