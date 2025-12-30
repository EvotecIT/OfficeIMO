using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Resizes shapes using the specified reference strategy.
        /// </summary>
        public void ResizeShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeSizeDimension dimension,
            PowerPointShapeSizeReference reference = PowerPointShapeSizeReference.First) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
            }

            (long width, long height) = ResolveReferenceSize(list, reference);

            switch (dimension) {
                case PowerPointShapeSizeDimension.Width:
                    foreach (PowerPointShape shape in list) {
                        shape.Width = width;
                    }
                    break;
                case PowerPointShapeSizeDimension.Height:
                    foreach (PowerPointShape shape in list) {
                        shape.Height = height;
                    }
                    break;
                case PowerPointShapeSizeDimension.Both:
                    foreach (PowerPointShape shape in list) {
                        shape.Width = width;
                        shape.Height = height;
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(dimension));
            }
        }

        /// <summary>
        ///     Resizes shapes to an explicit width/height (EMUs).
        /// </summary>
        public void ResizeShapes(IEnumerable<PowerPointShape> shapes, long? width, long? height) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
            }

            foreach (PowerPointShape shape in list) {
                if (width != null) {
                    shape.Width = width.Value;
                }
                if (height != null) {
                    shape.Height = height.Value;
                }
            }
        }

        /// <summary>
        ///     Resizes shapes to an explicit width/height (centimeters).
        /// </summary>
        public void ResizeShapesCm(IEnumerable<PowerPointShape> shapes, double? widthCm, double? heightCm) {
            long? width = widthCm != null ? PowerPointUnits.FromCentimeters(widthCm.Value) : null;
            long? height = heightCm != null ? PowerPointUnits.FromCentimeters(heightCm.Value) : null;
            ResizeShapes(shapes, width, height);
        }

        /// <summary>
        ///     Resizes shapes to an explicit width/height (inches).
        /// </summary>
        public void ResizeShapesInches(IEnumerable<PowerPointShape> shapes, double? widthInches, double? heightInches) {
            long? width = widthInches != null ? PowerPointUnits.FromInches(widthInches.Value) : null;
            long? height = heightInches != null ? PowerPointUnits.FromInches(heightInches.Value) : null;
            ResizeShapes(shapes, width, height);
        }

        /// <summary>
        ///     Resizes shapes to an explicit width/height (points).
        /// </summary>
        public void ResizeShapesPoints(IEnumerable<PowerPointShape> shapes, double? widthPoints, double? heightPoints) {
            long? width = widthPoints != null ? PowerPointUnits.FromPoints(widthPoints.Value) : null;
            long? height = heightPoints != null ? PowerPointUnits.FromPoints(heightPoints.Value) : null;
            ResizeShapes(shapes, width, height);
        }

        private static (long width, long height) ResolveReferenceSize(IReadOnlyList<PowerPointShape> shapes,
            PowerPointShapeSizeReference reference) {
            switch (reference) {
                case PowerPointShapeSizeReference.First:
                    PowerPointShape first = shapes[0];
                    return (first.Width, first.Height);
                case PowerPointShapeSizeReference.Smallest:
                    return (shapes.Min(s => s.Width), shapes.Min(s => s.Height));
                case PowerPointShapeSizeReference.Largest:
                    return (shapes.Max(s => s.Width), shapes.Max(s => s.Height));
                case PowerPointShapeSizeReference.Average:
                    long width = (long)Math.Round(shapes.Average(s => (double)s.Width));
                    long height = (long)Math.Round(shapes.Average(s => (double)s.Height));
                    return (width, height);
                default:
                    throw new ArgumentOutOfRangeException(nameof(reference));
            }
        }
    }
}
