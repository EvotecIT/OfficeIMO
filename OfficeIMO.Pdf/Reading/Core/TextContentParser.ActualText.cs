using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class TextContentParser {
    private static bool HasLogicalBoundsHint(object? propertyObject) =>
        propertyObject is PdfContentDictionary dictionary &&
        dictionary.Items.TryGetValue("OfficeIMOLogicalBounds", out object? value) && value is bool flag && flag;

    internal sealed partial class MarkedContentState {
        private readonly bool _allowExactGlyphMapping;
        internal bool PreserveAnchorGeometry { get; }

        /// <summary>Remembers the logical span without changing its painted selection geometry.</summary>
        internal void RememberLogicalLayoutSpan(List<PdfTextSpan> owner, int index) {
            RememberLogicalBaselineSpan(owner, index);
            if (PreserveAnchorGeometry || L._logicalPaintRejected || L._logicalLayoutOwner != null) return;
            L._logicalLayoutOwner = owner;
            L._logicalLayoutIndex = index;
            L._logicalLayoutText = owner[index].SourceActualText;
        }

        internal void RejectLogicalPaintExtent() => L._logicalPaintRejected = true;

        /// <summary>
        /// Accumulates a complete forward, same-baseline replacement from already decoded paint.
        /// Ambiguous paint discards the entire candidate, rather than publishing a partial width.
        /// </summary>
        internal void ObserveLogicalPaint(List<PdfTextSpan> owner, int? stream, string font, double size,
            double x, double y, double endX, double endY, Matrix2D transform, PdfPageClipPath? clip, double pageHeight,
            string paintedText, bool incompatiblePaint, bool invisibleAnchor, int lineBreaks) {
            if (PreserveAnchorGeometry || L._logicalPaintRejected) return;
            if (incompatiblePaint) { RejectLogicalPaintExtent(); return; }
            // The writer's initial and terminal semantic spaces do not paint a glyph.
            if (invisibleAnchor) return;
            if (!ReferenceEquals(owner, L._logicalLayoutOwner) || L._logicalLayoutText == null ||
                paintedText.Length == 0 || size <= 0D || !IsFinite(size) ||
                !IsFinite(x) || !IsFinite(y) || !IsFinite(endX) || !IsFinite(endY) ||
                !ContainsBaselineSegment(clip, x, y, endX, endY, pageHeight)) {
                RejectLogicalPaintExtent();
                return;
            }
            if (L._logicalPaintCharacters > L._logicalLayoutText.Length - paintedText.Length ||
                string.CompareOrdinal(L._logicalLayoutText, L._logicalPaintCharacters, paintedText, 0, paintedText.Length) != 0) {
                RejectLogicalPaintExtent();
                return;
            }
            if (!L._hasLogicalPaint) {
                PdfTextSpan first = owner[L._logicalLayoutIndex];
                double alongLength = Math.Sqrt(transform.A * transform.A + transform.B * transform.B);
                if (!first.IsVisible || first.ContentStreamObjectNumber != stream ||
                    Math.Abs(first.X - x) > 0.0001D || Math.Abs(first.Y - y) > 0.0001D ||
                    !(alongLength > 0D) || !IsFinite(alongLength) || !IsFinite(transform.C) || !IsFinite(transform.D) ||
                    transform.A * transform.D - transform.B * transform.C == 0D) {
                    RejectLogicalPaintExtent();
                    return;
                }
                L._hasLogicalPaint = true;
                L._logicalPaintX = x; L._logicalPaintY = y;
                L._logicalPaintAlongX = transform.A / alongLength;
                L._logicalPaintAlongY = transform.B / alongLength;
                L._logicalPaintFont = font; L._logicalPaintSize = size; L._logicalPaintStream = stream;
                L._logicalPaintTransform = transform; L._logicalPaintClip = clip;
            } else if (stream != L._logicalPaintStream || font != L._logicalPaintFont ||
                Math.Abs(size - L._logicalPaintSize) > 0.000001D || lineBreaks > 0 ||
                !SameLinearTransform(transform, L._logicalPaintTransform) || !CompatibleClipFrames(clip, L._logicalPaintClip)) {
                RejectLogicalPaintExtent();
                return;
            }
            double along = (x - L._logicalPaintX) * L._logicalPaintAlongX + (y - L._logicalPaintY) * L._logicalPaintAlongY;
            double normal = -(x - L._logicalPaintX) * L._logicalPaintAlongY + (y - L._logicalPaintY) * L._logicalPaintAlongX;
            double endAlong = (endX - L._logicalPaintX) * L._logicalPaintAlongX + (endY - L._logicalPaintY) * L._logicalPaintAlongY;
            double endNormal = -(endX - L._logicalPaintX) * L._logicalPaintAlongY + (endY - L._logicalPaintY) * L._logicalPaintAlongX;
            if (Math.Abs(normal) > 0.0001D || Math.Abs(endNormal) > 0.0001D ||
                along < L._logicalPaintLastOrigin - 0.0001D || endAlong < along ||
                !IsFinite(along) || !IsFinite(normal) || !IsFinite(endAlong) || !IsFinite(endNormal)) {
                RejectLogicalPaintExtent();
                return;
            }
            L._logicalPaintCharacters += paintedText.Length;
            L._logicalPaintLastOrigin = along;
            L._logicalPaintExtent = Math.Max(L._logicalPaintExtent, endAlong);
        }

        /// <summary>Commits only a closed, completely mapped replacement in its original stream.</summary>
        internal void CompleteLogicalPaintExtent(List<PdfTextSpan> owner) {
            if (_layout is null) return;
            if (!L._logicalPaintRejected && L._hasLogicalPaint && ReferenceEquals(owner, L._logicalLayoutOwner) &&
                L._logicalLayoutText != null && L._logicalPaintCharacters == L._logicalLayoutText.Length) {
                PdfTextSpan span = owner[L._logicalLayoutIndex];
                if (L._logicalPaintExtent > span.Advance) span.SetLogicalLayoutAdvance(L._logicalPaintExtent);
                // A closed inline scope can repeat an already mapped ligature without replacing it.
                // Extra shows, nested replacements and different logical text reject this proof above.
                if (_allowExactGlyphMapping && span.Text.Length > 0 &&
                    string.Equals(span.SourceActualText, span.Text, StringComparison.Ordinal) &&
                    string.Equals(span.RestampText, span.Text, StringComparison.Ordinal)) {
                    span.MarkExactActualTextGlyphMapping();
                }
            }
        }

        private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

        private static bool SameLinearTransform(Matrix2D first, Matrix2D second) =>
            Math.Abs(first.A - second.A) <= 0.000001D && Math.Abs(first.B - second.B) <= 0.000001D &&
            Math.Abs(first.C - second.C) <= 0.000001D && Math.Abs(first.D - second.D) <= 0.000001D;

        private static bool CompatibleClipFrames(PdfPageClipPath? first, PdfPageClipPath? second) {
            if (!first.HasValue || !second.HasValue) return first.HasValue == second.HasValue;
            PdfPageClipPath left = first.Value, right = second.Value;
            if (left.IsRectangle && right.IsRectangle) return true;
            // Repeated native frames can reconstruct the same rotated rectangle as
            // separate command lists. Compare exactly five commands, never arbitrary paths.
            if (left.X != right.X || left.Y != right.Y || left.Width != right.Width || left.Height != right.Height ||
                left.FillRule != right.FillRule || !IsFourCornerRectangle(left) || !IsFourCornerRectangle(right)) return false;
            for (int index = 0; index < 5; index++) {
                if (!left.Commands[index].Equals(right.Commands[index])) return false;
            }
            return true;
        }

        /// <summary>Admits complete baseline segments only through ordinary exact rectangle clips.</summary>
        private static bool ContainsBaselineSegment(PdfPageClipPath? clip, double x, double y,
            double endX, double endY, double pageHeight) {
            if (!clip.HasValue) return true;
            PdfPageClipPath bounds = clip.Value;
            if (!bounds.IsExact || bounds.ContainsTextClipping || !IsFinite(pageHeight)) return false;
            if (bounds.IsRectangle) {
                return Contains(x, pageHeight - y) && Contains(endX, pageHeight - endY);
                bool Contains(double px, double py) => px >= bounds.X - 0.0001D &&
                    px <= bounds.X + bounds.Width + 0.0001D && py >= bounds.Y - 0.0001D &&
                    py <= bounds.Y + bounds.Height + 0.0001D;
            }
            return IsFourCornerRectangle(bounds) &&
                ContainsRectanglePoint(bounds, x, pageHeight - y) && ContainsRectanglePoint(bounds, endX, pageHeight - endY);
        }

        private static bool IsFourCornerRectangle(PdfPageClipPath clip) {
            IReadOnlyList<OfficePathCommand> commands = clip.Commands;
            if (!clip.IsExact || clip.ContainsTextClipping || commands.Count != 5 ||
                commands[0].Kind != OfficePathCommandKind.MoveTo || commands[1].Kind != OfficePathCommandKind.LineTo ||
                commands[2].Kind != OfficePathCommandKind.LineTo || commands[3].Kind != OfficePathCommandKind.LineTo ||
                commands[4].Kind != OfficePathCommandKind.Close) return false;
            OfficePoint first = commands[0].Point, second = commands[1].Point,
                third = commands[2].Point, fourth = commands[3].Point;
            double ux = second.X - first.X, uy = second.Y - first.Y;
            double vx = third.X - second.X, vy = third.Y - second.Y;
            double lengths = Math.Sqrt((ux * ux + uy * uy) * (vx * vx + vy * vy));
            // Matrix serialization can slightly perturb orthogonality. Opposite edges
            // must still agree, and equality between two accepted frames remains exact.
            return lengths > 0D && IsFinite(lengths) && Math.Abs(ux * vx + uy * vy) <= lengths * 0.000001D &&
                Math.Abs(third.X - fourth.X - ux) <= 0.000001D && Math.Abs(third.Y - fourth.Y - uy) <= 0.000001D &&
                Math.Abs(fourth.X - first.X - vx) <= 0.000001D && Math.Abs(fourth.Y - first.Y - vy) <= 0.000001D;
        }

        private static bool ContainsRectanglePoint(PdfPageClipPath clip, double x, double y) {
            double sign = 0D;
            for (int index = 0; index < 4; index++) {
                OfficePoint first = clip.Commands[index].Point, second = clip.Commands[(index + 1) % 4].Point;
                double ux = second.X - first.X, uy = second.Y - first.Y;
                double cross = ux * (y - first.Y) - uy * (x - first.X);
                if (Math.Abs(cross) <= Math.Sqrt(ux * ux + uy * uy) * 0.0001D) continue;
                double edgeSign = Math.Sign(cross);
                if (sign != 0D && edgeSign != sign) return false;
                sign = edgeSign;
            }
            return true;
        }

        internal void RememberInvisibleAnchor(List<PdfTextSpan> owner, int index) {
            if (PreserveAnchorGeometry) return;
            L._invisibleAnchorOwner = owner;
            L._invisibleAnchorIndex = index;
        }

        internal bool TryGetInvisibleAnchorText(List<PdfTextSpan> owner, out string? text) {
            // Form parsing shares replacement ownership, but its span list has a
            // separate lifetime. Never replace an already returned parent stream.
            if (!ReferenceEquals(owner, L._invisibleAnchorOwner)) {
                text = null;
                return false;
            }
            PdfTextSpan anchor = owner[L._invisibleAnchorIndex];
            text = anchor.SourceActualText ?? anchor.Text;
            return true;
        }

        internal void ReplaceInvisibleAnchor(List<PdfTextSpan> owner, PdfTextSpan paint) {
            if (!ReferenceEquals(owner, L._invisibleAnchorOwner))
                throw new InvalidOperationException("ActualText geometry must belong to the current content stream.");
            // Geometry promotion spans multiple text objects and keeps a separate
            // semantic carrier. Removing just the paint object cannot safely edit
            // that replacement or preserve graphics state consumed later.
            owner[L._invisibleAnchorIndex] = paint.WithCanRestamp(false);
            L._invisibleAnchorOwner = null;
        }
    }
}
