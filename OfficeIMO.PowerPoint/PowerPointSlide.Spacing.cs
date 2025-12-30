using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Distributes shapes with a fixed spacing within their selection bounds.
        /// </summary>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            long spacingEmus, bool center = false) {
            DistributeShapesWithSpacing(shapes, distribution, GetSelectionBounds(shapes), spacingEmus, center);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within their selection bounds using alignment.
        /// </summary>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            long spacingEmus, PowerPointShapeAlignment alignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSelectionBounds(shapes), spacingEmus, alignment);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within their selection bounds using options.
        /// </summary>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointShapeSpacingOptions options) {
            DistributeShapesWithSpacing(shapes, distribution, GetSelectionBounds(shapes), options);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within their selection bounds using alignment and cross-axis alignment.
        /// </summary>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            long spacingEmus, PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSelectionBounds(shapes), spacingEmus, alignment, crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within a custom bounding box.
        /// </summary>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, long spacingEmus, bool center = false) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }
            if (spacingEmus < 0) {
                throw new ArgumentOutOfRangeException(nameof(spacingEmus));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count < 2) {
                return;
            }

            switch (distribution) {
                case PowerPointShapeDistribution.Horizontal:
                    DistributeHorizontalWithSpacing(list, bounds, spacingEmus, center);
                    break;
                case PowerPointShapeDistribution.Vertical:
                    DistributeVerticalWithSpacing(list, bounds, spacingEmus, center);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(distribution));
            }
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within a custom bounding box using alignment.
        /// </summary>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, long spacingEmus, PowerPointShapeAlignment alignment) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }
            if (spacingEmus < 0) {
                throw new ArgumentOutOfRangeException(nameof(spacingEmus));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count < 2) {
                return;
            }

            switch (distribution) {
                case PowerPointShapeDistribution.Horizontal:
                    DistributeHorizontalWithSpacingAligned(list, bounds, spacingEmus, alignment);
                    break;
                case PowerPointShapeDistribution.Vertical:
                    DistributeVerticalWithSpacingAligned(list, bounds, spacingEmus, alignment);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(distribution));
            }
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within a custom bounding box using options.
        /// </summary>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, PowerPointShapeSpacingOptions options) {
            PowerPointShapeSpacingOptions resolvedOptions = options ?? new PowerPointShapeSpacingOptions();
            PowerPointShapeAlignment alignment = resolvedOptions.Alignment ?? GetDefaultSpacingAlignment(distribution);
            DistributeShapesWithSpacingInternal(shapes, distribution, bounds, resolvedOptions.SpacingEmus, alignment,
                resolvedOptions.CrossAxisAlignment, resolvedOptions.ClampSpacingToBounds,
                resolvedOptions.ScaleToFitBounds, resolvedOptions.PreserveAspect);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within a custom bounding box using alignment and cross-axis alignment.
        /// </summary>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, long spacingEmus, PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }
            if (spacingEmus < 0) {
                throw new ArgumentOutOfRangeException(nameof(spacingEmus));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count < 2) {
                return;
            }

            switch (distribution) {
                case PowerPointShapeDistribution.Horizontal:
                    DistributeHorizontalWithSpacingAligned(list, bounds, spacingEmus, alignment, crossAxisAlignment);
                    break;
                case PowerPointShapeDistribution.Vertical:
                    DistributeVerticalWithSpacingAligned(list, bounds, spacingEmus, alignment, crossAxisAlignment);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(distribution));
            }
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the full slide bounds.
        /// </summary>
        public void DistributeShapesWithSpacingToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            long spacingEmus, bool center = false) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideBounds(), spacingEmus, center);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the full slide bounds using alignment.
        /// </summary>
        public void DistributeShapesWithSpacingToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            long spacingEmus, PowerPointShapeAlignment alignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideBounds(), spacingEmus, alignment);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the full slide bounds using options.
        /// </summary>
        public void DistributeShapesWithSpacingToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointShapeSpacingOptions options) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideBounds(), options);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the full slide bounds using alignment and cross-axis alignment.
        /// </summary>
        public void DistributeShapesWithSpacingToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            long spacingEmus, PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideBounds(), spacingEmus, alignment, crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the slide content bounds using a margin (EMUs).
        /// </summary>
        public void DistributeShapesWithSpacingToSlideContent(IEnumerable<PowerPointShape> shapes,
            PowerPointShapeDistribution distribution, long spacingEmus, long marginEmus, bool center = false) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideContentBounds(marginEmus), spacingEmus, center);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the slide content bounds using a margin (EMUs) and alignment.
        /// </summary>
        public void DistributeShapesWithSpacingToSlideContent(IEnumerable<PowerPointShape> shapes,
            PowerPointShapeDistribution distribution, long spacingEmus, long marginEmus, PowerPointShapeAlignment alignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideContentBounds(marginEmus), spacingEmus, alignment);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the slide content bounds using a margin (EMUs) and options.
        /// </summary>
        public void DistributeShapesWithSpacingToSlideContent(IEnumerable<PowerPointShape> shapes,
            PowerPointShapeDistribution distribution, long marginEmus, PowerPointShapeSpacingOptions options) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideContentBounds(marginEmus), options);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the slide content bounds using a margin (EMUs),
        ///     alignment, and cross-axis alignment.
        /// </summary>
        public void DistributeShapesWithSpacingToSlideContent(IEnumerable<PowerPointShape> shapes,
            PowerPointShapeDistribution distribution, long spacingEmus, long marginEmus,
            PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideContentBounds(marginEmus), spacingEmus,
                alignment, crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the slide content bounds using a margin (centimeters).
        /// </summary>
        public void DistributeShapesWithSpacingToSlideContentCm(IEnumerable<PowerPointShape> shapes,
            PowerPointShapeDistribution distribution, double spacingCm, double marginCm, bool center = false) {
            DistributeShapesWithSpacingToSlideContent(shapes, distribution,
                PowerPointUnits.FromCentimeters(spacingCm),
                PowerPointUnits.FromCentimeters(marginCm),
                center);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the slide content bounds using a margin (inches).
        /// </summary>
        public void DistributeShapesWithSpacingToSlideContentInches(IEnumerable<PowerPointShape> shapes,
            PowerPointShapeDistribution distribution, double spacingInches, double marginInches, bool center = false) {
            DistributeShapesWithSpacingToSlideContent(shapes, distribution,
                PowerPointUnits.FromInches(spacingInches),
                PowerPointUnits.FromInches(marginInches),
                center);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the slide content bounds using a margin (points).
        /// </summary>
        public void DistributeShapesWithSpacingToSlideContentPoints(IEnumerable<PowerPointShape> shapes,
            PowerPointShapeDistribution distribution, double spacingPoints, double marginPoints, bool center = false) {
            DistributeShapesWithSpacingToSlideContent(shapes, distribution,
                PowerPointUnits.FromPoints(spacingPoints),
                PowerPointUnits.FromPoints(marginPoints),
                center);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing using centimeters.
        /// </summary>
        public void DistributeShapesWithSpacingCm(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            double spacingCm, bool center = false) {
            DistributeShapesWithSpacing(shapes, distribution, PowerPointUnits.FromCentimeters(spacingCm), center);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing using inches.
        /// </summary>
        public void DistributeShapesWithSpacingInches(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            double spacingInches, bool center = false) {
            DistributeShapesWithSpacing(shapes, distribution, PowerPointUnits.FromInches(spacingInches), center);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing using points.
        /// </summary>
        public void DistributeShapesWithSpacingPoints(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            double spacingPoints, bool center = false) {
            DistributeShapesWithSpacing(shapes, distribution, PowerPointUnits.FromPoints(spacingPoints), center);
        }

        private static void DistributeHorizontalWithSpacing(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, bool center) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Left).ToList();
            long totalWidth = ordered.Sum(s => s.Width);
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockWidth = totalWidth + totalSpacing;
            double start = bounds.Left;

            if (center && blockWidth < bounds.Width) {
                start = bounds.Left + (bounds.Width - blockWidth) / 2d;
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Left = (long)Math.Round(current);
                current += shape.Width + spacingEmus;
            }
        }

        private static void DistributeVerticalWithSpacing(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, bool center) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Top).ToList();
            long totalHeight = ordered.Sum(s => s.Height);
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockHeight = totalHeight + totalSpacing;
            double start = bounds.Top;

            if (center && blockHeight < bounds.Height) {
                start = bounds.Top + (bounds.Height - blockHeight) / 2d;
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Top = (long)Math.Round(current);
                current += shape.Height + spacingEmus;
            }
        }

        private static void DistributeHorizontalWithSpacingAligned(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Left).ToList();
            long totalWidth = ordered.Sum(s => s.Width);
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockWidth = totalWidth + totalSpacing;

            double start = bounds.Left;
            if (blockWidth <= bounds.Width) {
                start = alignment switch {
                    PowerPointShapeAlignment.Left => bounds.Left,
                    PowerPointShapeAlignment.Center => bounds.Left + (bounds.Width - blockWidth) / 2d,
                    PowerPointShapeAlignment.Right => bounds.Right - blockWidth,
                    _ => throw new ArgumentOutOfRangeException(nameof(alignment))
                };
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Left = (long)Math.Round(current);
                current += shape.Width + spacingEmus;
            }
        }

        private static void DistributeHorizontalWithSpacingAligned(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Left).ToList();
            long totalWidth = ordered.Sum(s => s.Width);
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockWidth = totalWidth + totalSpacing;

            double start = bounds.Left;
            if (blockWidth <= bounds.Width) {
                start = alignment switch {
                    PowerPointShapeAlignment.Left => bounds.Left,
                    PowerPointShapeAlignment.Center => bounds.Left + (bounds.Width - blockWidth) / 2d,
                    PowerPointShapeAlignment.Right => bounds.Right - blockWidth,
                    _ => throw new ArgumentOutOfRangeException(nameof(alignment))
                };
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Left = (long)Math.Round(current);
                current += shape.Width + spacingEmus;
            }

            ApplyCrossAxisAlignment(ordered, PowerPointShapeDistribution.Horizontal, bounds, crossAxisAlignment);
        }

        private static void DistributeVerticalWithSpacingAligned(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Top).ToList();
            long totalHeight = ordered.Sum(s => s.Height);
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockHeight = totalHeight + totalSpacing;

            double start = bounds.Top;
            if (blockHeight <= bounds.Height) {
                start = alignment switch {
                    PowerPointShapeAlignment.Top => bounds.Top,
                    PowerPointShapeAlignment.Middle => bounds.Top + (bounds.Height - blockHeight) / 2d,
                    PowerPointShapeAlignment.Bottom => bounds.Bottom - blockHeight,
                    _ => throw new ArgumentOutOfRangeException(nameof(alignment))
                };
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Top = (long)Math.Round(current);
                current += shape.Height + spacingEmus;
            }
        }

        private void DistributeShapesWithSpacingInternal(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, long spacingEmus, PowerPointShapeAlignment alignment,
            PowerPointShapeAlignment? crossAxisAlignment, bool clampSpacingToBounds,
            bool scaleToFitBounds, bool preserveAspect) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }
            if (spacingEmus < 0) {
                throw new ArgumentOutOfRangeException(nameof(spacingEmus));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count < 2) {
                return;
            }

            long effectiveSpacing = ResolveSpacingForBounds(list, distribution, bounds, spacingEmus, clampSpacingToBounds);
            if (scaleToFitBounds) {
                ApplyScaleToFitBounds(list, distribution, bounds, effectiveSpacing, preserveAspect);
            }

            switch (distribution) {
                case PowerPointShapeDistribution.Horizontal:
                    if (crossAxisAlignment.HasValue) {
                        DistributeHorizontalWithSpacingAligned(list, bounds, effectiveSpacing, alignment, crossAxisAlignment.Value);
                    } else {
                        DistributeHorizontalWithSpacingAligned(list, bounds, effectiveSpacing, alignment);
                    }
                    break;
                case PowerPointShapeDistribution.Vertical:
                    if (crossAxisAlignment.HasValue) {
                        DistributeVerticalWithSpacingAligned(list, bounds, effectiveSpacing, alignment, crossAxisAlignment.Value);
                    } else {
                        DistributeVerticalWithSpacingAligned(list, bounds, effectiveSpacing, alignment);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(distribution));
            }
        }

        private static long ResolveSpacingForBounds(IReadOnlyList<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, long spacingEmus, bool clampSpacingToBounds) {
            if (!clampSpacingToBounds || shapes.Count < 2) {
                return spacingEmus;
            }

            long totalSize = 0;
            for (int i = 0; i < shapes.Count; i++) {
                totalSize += distribution == PowerPointShapeDistribution.Horizontal
                    ? shapes[i].Width
                    : shapes[i].Height;
            }

            long available = (distribution == PowerPointShapeDistribution.Horizontal ? bounds.Width : bounds.Height) - totalSize;
            if (available <= 0) {
                return 0;
            }

            long maxSpacing = (long)Math.Floor(available / (double)(shapes.Count - 1));
            if (maxSpacing < 0) {
                return 0;
            }

            return Math.Min(spacingEmus, maxSpacing);
        }

        private static PowerPointShapeAlignment GetDefaultSpacingAlignment(PowerPointShapeDistribution distribution) {
            return distribution == PowerPointShapeDistribution.Horizontal
                ? PowerPointShapeAlignment.Left
                : PowerPointShapeAlignment.Top;
        }

        private static void ApplyScaleToFitBounds(IReadOnlyList<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, long spacingEmus, bool preserveAspect) {
            if (shapes.Count == 0) {
                return;
            }

            long totalSize = 0;
            for (int i = 0; i < shapes.Count; i++) {
                totalSize += distribution == PowerPointShapeDistribution.Horizontal
                    ? shapes[i].Width
                    : shapes[i].Height;
            }

            long available = (distribution == PowerPointShapeDistribution.Horizontal ? bounds.Width : bounds.Height)
                - spacingEmus * Math.Max(0, shapes.Count - 1);
            if (totalSize <= 0 || available >= totalSize) {
                return;
            }

            double scale = available <= 0 ? 0d : available / (double)totalSize;
            ScaleShapesForAxis(shapes, distribution == PowerPointShapeDistribution.Horizontal, scale, preserveAspect);
        }

        private static void DistributeVerticalWithSpacingAligned(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Top).ToList();
            long totalHeight = ordered.Sum(s => s.Height);
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockHeight = totalHeight + totalSpacing;

            double start = bounds.Top;
            if (blockHeight <= bounds.Height) {
                start = alignment switch {
                    PowerPointShapeAlignment.Top => bounds.Top,
                    PowerPointShapeAlignment.Middle => bounds.Top + (bounds.Height - blockHeight) / 2d,
                    PowerPointShapeAlignment.Bottom => bounds.Bottom - blockHeight,
                    _ => throw new ArgumentOutOfRangeException(nameof(alignment))
                };
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Top = (long)Math.Round(current);
                current += shape.Height + spacingEmus;
            }

            ApplyCrossAxisAlignment(ordered, PowerPointShapeDistribution.Vertical, bounds, crossAxisAlignment);
        }
    }
}
