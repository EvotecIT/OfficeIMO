using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class TextContentParser {
    internal sealed partial class MarkedContentState {
        private List<PdfTextSpan>? _logicalBaselineOwner;
        private int _logicalBaselineIndex;
        private bool _logicalBaselineRejected, _logicalBaselineStarted, _logicalBaselineHasBase;
        private int? _logicalBaselineStream;
        private string? _logicalBaselineFont;
        private double _logicalBaselineSize, _logicalBaselineX, _logicalBaselineY;
        private double _logicalBaselineAlongX, _logicalBaselineAlongY;
        private double _logicalBaselineBaseNormal, _logicalBaselineLastBaseOrigin;
        private Matrix2D _logicalBaselineTransform;
        private PdfPageClipPath? _logicalBaselineClip;

        /// <summary>Tracks a replacement separately from its first glyph's painted origin.</summary>
        internal void RememberLogicalBaselineSpan(List<PdfTextSpan> owner, int index) {
            if (PreserveAnchorGeometry || _logicalBaselineRejected || _logicalBaselineOwner != null) return;
            _logicalBaselineOwner = owner;
            _logicalBaselineIndex = index;
        }

        /// <summary>Invalidates both independent layout proofs for nested scopes or Form paint.</summary>
        internal void RejectLogicalLayoutEvidence() {
            RejectLogicalPaintExtent();
            _logicalBaselineRejected = true;
        }

        /// <summary>
        /// Finds a common advancing baseline after an initial zero-advance mark. ActualText may
        /// replace substituted glyph text, so this geometric proof does not claim source mapping.
        /// </summary>
        internal void ObserveLogicalBaselinePaint(List<PdfTextSpan> owner, int? stream, string font, double size,
            double x, double y, double endX, double endY, Matrix2D transform, PdfPageClipPath? clip,
            double pageHeight, string paintedText, bool incompatiblePaint, bool invisibleAnchor, int lineBreaks) {
            if (PreserveAnchorGeometry || _logicalBaselineRejected) return;
            if (invisibleAnchor || incompatiblePaint || !ReferenceEquals(owner, _logicalBaselineOwner) ||
                paintedText.Length == 0 || size <= 0D || !IsFinite(size) ||
                !IsFinite(x) || !IsFinite(y) || !IsFinite(endX) || !IsFinite(endY) ||
                !IsFinite(transform.E) || !IsFinite(transform.F) ||
                !ContainsBaselineSegment(clip, x, y, endX, endY, pageHeight)) {
                _logicalBaselineRejected = true;
                return;
            }
            if (!_logicalBaselineStarted) {
                PdfTextSpan first = owner[_logicalBaselineIndex];
                double length = Math.Sqrt(transform.A * transform.A + transform.B * transform.B);
                if (!first.IsVisible || first.ContentStreamObjectNumber != stream ||
                    Math.Abs(first.Advance) > 0.0001D ||
                    !OfficeTextElements.ContainsCombiningMark(paintedText) ||
                    !(length > 0D) || !IsFinite(length) || !IsFinite(transform.C) || !IsFinite(transform.D) ||
                    transform.A * transform.D - transform.B * transform.C == 0D) {
                    _logicalBaselineRejected = true;
                    return;
                }
                _logicalBaselineStarted = true;
                _logicalBaselineStream = stream; _logicalBaselineFont = font; _logicalBaselineSize = size;
                _logicalBaselineTransform = transform; _logicalBaselineClip = clip;
                _logicalBaselineAlongX = transform.A / length; _logicalBaselineAlongY = transform.B / length;
                _logicalBaselineX = transform.E; _logicalBaselineY = transform.F;
            } else if (stream != _logicalBaselineStream || font != _logicalBaselineFont ||
                Math.Abs(size - _logicalBaselineSize) > 0.000001D || lineBreaks > 0 ||
                !SameLinearTransform(transform, _logicalBaselineTransform) || !SameLogicalBaselineClip(clip)) {
                _logicalBaselineRejected = true;
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
                _logicalBaselineRejected = true;
                return;
            }
            if (endAlong - along <= 0.0001D) return;
            if (_logicalBaselineHasBase && (Math.Abs(normal - _logicalBaselineBaseNormal) > 0.0001D ||
                along < _logicalBaselineLastBaseOrigin - 0.0001D)) {
                _logicalBaselineRejected = true;
                return;
            }
            _logicalBaselineHasBase = true;
            _logicalBaselineBaseNormal = normal;
            _logicalBaselineLastBaseOrigin = along;
        }

        /// <summary>Commits only closed, compatible mark-first paint; public glyph geometry is unchanged.</summary>
        internal void CompleteLogicalBaseline(List<PdfTextSpan> owner) {
            if (_logicalBaselineRejected || !_logicalBaselineHasBase ||
                !ReferenceEquals(owner, _logicalBaselineOwner)) return;
            PdfTextSpan span = owner[_logicalBaselineIndex];
            double offset = _logicalBaselineBaseNormal - Normal(span.X, span.Y);
            if (IsFinite(offset)) span.SetLogicalLayoutBaselineOffset(offset);
        }

        private double Along(double x, double y) =>
            (x - _logicalBaselineX) * _logicalBaselineAlongX + (y - _logicalBaselineY) * _logicalBaselineAlongY;

        private double Normal(double x, double y) =>
            -(x - _logicalBaselineX) * _logicalBaselineAlongY + (y - _logicalBaselineY) * _logicalBaselineAlongX;

        private bool SameLogicalBaselineClip(PdfPageClipPath? clip) {
            if (!clip.HasValue || !_logicalBaselineClip.HasValue) return clip.HasValue == _logicalBaselineClip.HasValue;
            PdfPageClipPath left = clip.Value, right = _logicalBaselineClip.Value;
            return left.X == right.X && left.Y == right.Y && left.Width == right.Width && left.Height == right.Height &&
                CompatibleClipFrames(clip, _logicalBaselineClip);
        }
    }
}
