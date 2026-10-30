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
                        shape.Width= OfficeIMO.OfficeLength.Emus(width);
                    }
                    break;
                case PowerPointShapeSizeDimension.Height:
                    foreach (PowerPointShape shape in list) {
                        shape.Height= OfficeIMO.OfficeLength.Emus(height);
                    }
                    break;
                case PowerPointShapeSizeDimension.Both:
                    foreach (PowerPointShape shape in list) {
                        shape.Width= OfficeIMO.OfficeLength.Emus(width);
                        shape.Height= OfficeIMO.OfficeLength.Emus(height);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(dimension));
            }
        }

        /// <summary>
        ///     Resizes shapes to explicit dimensions, leaving each null dimension unchanged.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="width">Nonnegative width, or null to retain each shape's current width. Zero is permitted.</param>
        /// <param name="height">Nonnegative height, or null to retain each shape's current height. Zero is permitted.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ResizeShapes(IEnumerable<PowerPointShape> shapes, OfficeLength? width, OfficeLength? height) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            if (width.HasValue && width.Value < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(width));
            }
            if (height.HasValue && height.Value < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(height));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
            }

            foreach (PowerPointShape shape in list) {
                if (width != null) {
                    shape.Width= width.Value;
                }
                if (height != null) {
                    shape.Height= height.Value;
                }
            }
        }

        private static (long width, long height) ResolveReferenceSize(IReadOnlyList<PowerPointShape> shapes,
            PowerPointShapeSizeReference reference) {
            switch (reference) {
                case PowerPointShapeSizeReference.First:
                    PowerPointShape first = shapes[0];
                    return (first.Width.ToEmus(), first.Height.ToEmus());
                case PowerPointShapeSizeReference.Smallest:
                    return (shapes.Min(s => s.Width.ToEmus()), shapes.Min(s => s.Height.ToEmus()));
                case PowerPointShapeSizeReference.Largest:
                    return (shapes.Max(s => s.Width.ToEmus()), shapes.Max(s => s.Height.ToEmus()));
                case PowerPointShapeSizeReference.Average:
                    long width = (long)Math.Round(shapes.Average(s => (double)s.Width.ToEmus()));
                    long height = (long)Math.Round(shapes.Average(s => (double)s.Height.ToEmus()));
                    return (width, height);
                default:
                    throw new ArgumentOutOfRangeException(nameof(reference));
            }
        }
    }
}
