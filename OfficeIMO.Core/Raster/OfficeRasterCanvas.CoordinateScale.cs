using System.Collections.Generic;

namespace OfficeIMO.Drawing {
    public sealed partial class OfficeRasterCanvas {
        // Raster layers can sample the two local axes independently. Layout and
        // contour construction retain their nominal units; only painting, clipping
        // and image projection cross this device-coordinate boundary.
        internal double CoordinateScaleX { get; private set; } = 1D;
        internal double CoordinateScaleY { get; private set; } = 1D;

        internal void SetCoordinateScale(double scaleX, double scaleY) {
            CoordinateScaleX = scaleX;
            CoordinateScaleY = scaleY;
        }

        private bool HasCoordinateScale => CoordinateScaleX != 1D || CoordinateScaleY != 1D;

        private OfficeTransform ScaleCoordinates(OfficeTransform transform) => HasCoordinateScale
            ? transform.Then(OfficeTransform.Scale(CoordinateScaleX, CoordinateScaleY)) : transform;

        private IReadOnlyList<OfficePoint> ScaleCoordinates(IReadOnlyList<OfficePoint> points) {
            if (!HasCoordinateScale) return points;
            OfficePoint[] result = new OfficePoint[points.Count];
            for (int index = 0; index < points.Count; index++) {
                result[index] = new OfficePoint(points[index].X * CoordinateScaleX, points[index].Y * CoordinateScaleY);
            }
            return result;
        }

        private IReadOnlyList<IReadOnlyList<OfficePoint>> ScaleCoordinates(IReadOnlyList<IReadOnlyList<OfficePoint>> contours) {
            if (!HasCoordinateScale) return contours;
            IReadOnlyList<OfficePoint>[] result = new IReadOnlyList<OfficePoint>[contours.Count];
            for (int index = 0; index < contours.Count; index++) result[index] = ScaleCoordinates(contours[index]);
            return result;
        }
    }
}
