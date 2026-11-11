using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class TextContentParser {
    internal sealed partial class MarkedContentState {

        /// <summary>Tracks a replacement separately from its first glyph's painted origin.</summary>
        internal void RememberLogicalBaselineSpan(List<PdfTextSpan> owner, int index) {
            if (PreserveAnchorGeometry || L._logicalBaselineRejected || L._logicalBaselineOwner != null) return;
            L._logicalBaselineOwner = owner;
            L._logicalBaselineIndex = index;
        }

        /// <summary>Invalidates both independent layout proofs for nested scopes or Form paint.</summary>
        internal void RejectLogicalLayoutEvidence() {
            RejectLogicalPaintExtent();
            L._logicalBaselineRejected = true;
        }

        /// <summary>
        /// Finds a common advancing baseline after an initial zero-advance mark. ActualText may
        /// replace substituted glyph text, so this geometric proof does not claim source mapping.
        /// </summary>
        internal void ObserveLogicalBaselinePaint(List<PdfTextSpan> owner, int? stream, string font, double size,
            double x, double y, double endX, double endY, Matrix2D transform, PdfPageClipPath? clip,
            double pageHeight, string paintedText, bool incompatiblePaint, bool invisibleAnchor, int lineBreaks) {
            if (PreserveAnchorGeometry || L._logicalBaselineRejected) return;
            if (invisibleAnchor || incompatiblePaint || !ReferenceEquals(owner, L._logicalBaselineOwner) ||
                paintedText.Length == 0 || size <= 0D || !IsFinite(size) ||
                !IsFinite(x) || !IsFinite(y) || !IsFinite(endX) || !IsFinite(endY) ||
                !IsFinite(transform.E) || !IsFinite(transform.F) ||
                !ContainsBaselineSegment(clip, x, y, endX, endY, pageHeight)) {
                L._logicalBaselineRejected = true;
                return;
            }
            if (!L._logicalBaselineStarted) {
                PdfTextSpan first = owner[L._logicalBaselineIndex];
                double length = Math.Sqrt(transform.A * transform.A + transform.B * transform.B);
                if (!first.IsVisible || first.ContentStreamObjectNumber != stream ||
                    Math.Abs(first.Advance) > 0.0001D ||
                    !OfficeTextElements.ContainsCombiningMark(paintedText) ||
                    !(length > 0D) || !IsFinite(length) || !IsFinite(transform.C) || !IsFinite(transform.D) ||
                    transform.A * transform.D - transform.B * transform.C == 0D) {
                    L._logicalBaselineRejected = true;
                    return;
                }
                L._logicalBaselineStarted = true;
                L._logicalBaselineStream = stream; L._logicalBaselineFont = font; L._logicalBaselineSize = size;
                L._logicalBaselineTransform = transform; L._logicalBaselineClip = clip;
                L._logicalBaselineAlongX = transform.A / length; L._logicalBaselineAlongY = transform.B / length;
                L._logicalBaselineX = transform.E; L._logicalBaselineY = transform.F;
            } else if (stream != L._logicalBaselineStream || font != L._logicalBaselineFont ||
                Math.Abs(size - L._logicalBaselineSize) > 0.000001D || lineBreaks > 0 ||
                !SameLinearTransform(transform, L._logicalBaselineTransform) || !SameLogicalBaselineClip(clip)) {
                L._logicalBaselineRejected = true;
                return;
            }
            // Ts may move marks across the baseline. The underlying text matrices must still
            // describe one line, and every advancing run must agree on its actual raised baseline.
            double matrixNormal = Normal(transform.E, transform.F);
            double along = Along(x, y), endAlong = Along(endX, endY);
            double normal = Normal(x, y), endNormal = Normal(endX, endY);
            if (Math.Abs(matrixNormal) > 0.0001D || !IsFinite(along) || !IsFinite(endAlong) ||
                !IsFinite(normal) || !IsFinite(endNormal) || Math.Abs(endNormal - normal) > 0.0001D ||
                endAlong < along - 0.0001D) {
                L._logicalBaselineRejected = true;
                return;
            }
            if (endAlong - along <= 0.0001D) return;
            if (L._logicalBaselineHasBase && (Math.Abs(normal - L._logicalBaselineBaseNormal) > 0.0001D ||
                along < L._logicalBaselineLastBaseOrigin - 0.0001D)) {
                L._logicalBaselineRejected = true;
                return;
            }
            L._logicalBaselineHasBase = true;
            L._logicalBaselineBaseNormal = normal;
            L._logicalBaselineLastBaseOrigin = along;
        }

        /// <summary>Commits only closed, compatible mark-first paint; public glyph geometry is unchanged.</summary>
        internal void CompleteLogicalBaseline(List<PdfTextSpan> owner) {
            if (_layout is null) return;
            if (L._logicalBaselineRejected || !L._logicalBaselineHasBase ||
                !ReferenceEquals(owner, L._logicalBaselineOwner)) return;
            PdfTextSpan span = owner[L._logicalBaselineIndex];
            double offset = L._logicalBaselineBaseNormal - Normal(span.X, span.Y);
            if (IsFinite(offset)) span.SetLogicalLayoutBaselineOffset(offset);
        }

        private double Along(double x, double y) =>
            (x - L._logicalBaselineX) * L._logicalBaselineAlongX + (y - L._logicalBaselineY) * L._logicalBaselineAlongY;

        private double Normal(double x, double y) =>
            -(x - L._logicalBaselineX) * L._logicalBaselineAlongY + (y - L._logicalBaselineY) * L._logicalBaselineAlongX;

        private bool SameLogicalBaselineClip(PdfPageClipPath? clip) {
            if (!clip.HasValue || !L._logicalBaselineClip.HasValue) return clip.HasValue == L._logicalBaselineClip.HasValue;
            PdfPageClipPath left = clip.Value, right = L._logicalBaselineClip.Value;
            return left.X == right.X && left.Y == right.Y && left.Width == right.Width && left.Height == right.Height &&
                CompatibleClipFrames(clip, L._logicalBaselineClip);
        }
    }
}
