using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private OfficeTransform? _positionedGlyphTransform;

    // Unlike a drawing/image transform, this changes glyph geometry before the
    // pen is applied. Text-matrix shear must not shear the width or dash pattern.
    internal IDisposable PushPositionedGlyphTransform(OfficeTransform transform) {
        var scope = new PositionedGlyphTransformScope(this, _positionedGlyphTransform);
        _positionedGlyphTransform = transform;
        return scope;
    }

    private sealed class PositionedGlyphTransformScope : IDisposable {
        private OfficeRasterCanvas? _canvas;
        private readonly OfficeTransform? _previous;
        internal PositionedGlyphTransformScope(OfficeRasterCanvas canvas, OfficeTransform? previous) {
            _canvas = canvas;
            _previous = previous;
        }
        public void Dispose() {
            if (_canvas == null) return;
            _canvas._positionedGlyphTransform = _previous;
            _canvas = null;
        }
    }

    private IReadOnlyList<List<OfficePoint>> TransformPositionedGlyphContours(
        IReadOnlyList<List<OfficePoint>> contours, double boldOffset = 0D) {
        if (!_positionedGlyphTransform.HasValue && boldOffset == 0D) return contours;
        var transformed = new List<List<OfficePoint>>(contours.Count);
        foreach (List<OfficePoint> contour in contours) {
            _cancellationToken.ThrowIfCancellationRequested();
            var points = new List<OfficePoint>(contour.Count);
            foreach (OfficePoint point in contour) {
                _cancellationToken.ThrowIfCancellationRequested();
                var shifted = new OfficePoint(point.X + boldOffset, point.Y);
                points.Add(_positionedGlyphTransform.HasValue
                    ? _positionedGlyphTransform.Value.TransformPoint(shifted) : shifted);
            }
            transformed.Add(points);
        }
        return transformed;
    }
}
