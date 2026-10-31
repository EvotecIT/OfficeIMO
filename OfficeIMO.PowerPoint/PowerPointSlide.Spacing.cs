using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Distributes shapes with a fixed spacing within their selection bounds.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="center">Whether to center the arranged shapes within the target bounds.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            OfficeLength spacing, bool center = false) {
            DistributeShapesWithSpacing(shapes, distribution, GetSelectionBounds(shapes), spacing, center);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within their selection bounds using alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            OfficeLength spacing, PowerPointShapeAlignment alignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSelectionBounds(shapes), spacing, alignment);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within their selection bounds using options.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointShapeSpacingOptions? options) {
            DistributeShapesWithSpacing(shapes, distribution, GetSelectionBounds(shapes), options);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within their selection bounds using alignment and cross-axis alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <param name="crossAxisAlignment">Alignment on the axis perpendicular to distribution.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            OfficeLength spacing, PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSelectionBounds(shapes), spacing, alignment, crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within a custom bounding box.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="center">Whether to center the arranged shapes within the target bounds.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, OfficeLength spacing, bool center = false) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }
            if (spacing < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(spacing));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count < 2) {
                return;
            }

            ValidateLayoutBounds(bounds);
            ValidateDistributionLayout(distribution);
            switch (distribution) {
                case PowerPointShapeDistribution.Horizontal:
                    DistributeHorizontalWithSpacing(list, bounds, spacing.ToEmus(), center);
                    break;
                case PowerPointShapeDistribution.Vertical:
                    DistributeVerticalWithSpacing(list, bounds, spacing.ToEmus(), center);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(distribution));
            }
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within a custom bounding box using alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, OfficeLength spacing, PowerPointShapeAlignment alignment) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }
            if (spacing < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(spacing));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count < 2) {
                return;
            }

            ValidateLayoutBounds(bounds);
            ValidateDistributionLayout(distribution, alignment);
            switch (distribution) {
                case PowerPointShapeDistribution.Horizontal:
                    DistributeHorizontalWithSpacingAligned(list, bounds, spacing.ToEmus(), alignment);
                    break;
                case PowerPointShapeDistribution.Vertical:
                    DistributeVerticalWithSpacingAligned(list, bounds, spacing.ToEmus(), alignment);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(distribution));
            }
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within a custom bounding box using options.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, PowerPointShapeSpacingOptions? options) {
            PowerPointShapeSpacingOptions resolvedOptions = options ?? new PowerPointShapeSpacingOptions();
            PowerPointShapeAlignment alignment = resolvedOptions.Alignment ?? GetDefaultSpacingAlignment(distribution);
            DistributeShapesWithSpacingInternal(shapes, distribution, bounds, resolvedOptions.Spacing.ToEmus(), alignment,
                resolvedOptions.CrossAxisAlignment, resolvedOptions.ClampSpacingToBounds,
                resolvedOptions.ScaleToFitBounds, resolvedOptions.PreserveAspect);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within a custom bounding box using alignment and cross-axis alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <param name="crossAxisAlignment">Alignment on the axis perpendicular to distribution.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacing(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, OfficeLength spacing, PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }
            if (spacing < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(spacing));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count < 2) {
                return;
            }

            ValidateLayoutBounds(bounds);
            ValidateDistributionLayout(distribution, alignment, crossAxisAlignment);
            switch (distribution) {
                case PowerPointShapeDistribution.Horizontal:
                    DistributeHorizontalWithSpacingAligned(list, bounds, spacing.ToEmus(), alignment, crossAxisAlignment);
                    break;
                case PowerPointShapeDistribution.Vertical:
                    DistributeVerticalWithSpacingAligned(list, bounds, spacing.ToEmus(), alignment, crossAxisAlignment);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(distribution));
            }
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the full slide bounds.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="center">Whether to center the arranged shapes within the target bounds.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacingToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            OfficeLength spacing, bool center = false) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideBounds(), spacing, center);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the full slide bounds using alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacingToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            OfficeLength spacing, PowerPointShapeAlignment alignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideBounds(), spacing, alignment);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the full slide bounds using options.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        public void DistributeShapesWithSpacingToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointShapeSpacingOptions? options) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideBounds(), options);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the full slide bounds using alignment and cross-axis alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <param name="crossAxisAlignment">Alignment on the axis perpendicular to distribution.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacingToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            OfficeLength spacing, PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideBounds(), spacing, alignment, crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the slide content bounds using a margin.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="center">Whether to center the arranged shapes within the target bounds.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacingToSlideContent(IEnumerable<PowerPointShape> shapes,
            PowerPointShapeDistribution distribution, OfficeLength spacing, OfficeLength margin, bool center = false) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideContentBounds(margin), spacing, center);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the slide content bounds using a margin and alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacingToSlideContent(IEnumerable<PowerPointShape> shapes,
            PowerPointShapeDistribution distribution, OfficeLength spacing, OfficeLength margin, PowerPointShapeAlignment alignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideContentBounds(margin), spacing, alignment);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the slide content bounds using a margin and options.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacingToSlideContent(IEnumerable<PowerPointShape> shapes,
            PowerPointShapeDistribution distribution, OfficeLength margin, PowerPointShapeSpacingOptions? options) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideContentBounds(margin), options);
        }

        /// <summary>
        ///     Distributes shapes with a fixed spacing within the slide content bounds using a margin,
        ///     alignment, and cross-axis alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <param name="crossAxisAlignment">Alignment on the axis perpendicular to distribution.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesWithSpacingToSlideContent(IEnumerable<PowerPointShape> shapes,
            PowerPointShapeDistribution distribution, OfficeLength spacing, OfficeLength margin,
            PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapesWithSpacing(shapes, distribution, GetSlideContentBounds(margin), spacing,
                alignment, crossAxisAlignment);
        }

        private static void DistributeHorizontalWithSpacing(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, bool center) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Left.ToEmus()).ToList();
            long totalWidth = ordered.Sum(s => s.Width.ToEmus());
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockWidth = totalWidth + totalSpacing;
            double start = bounds.Left.ToEmus();

            if (center && blockWidth < bounds.Width.ToEmus()) {
                start = bounds.Left.ToEmus() + (bounds.Width.ToEmus() - blockWidth) / 2d;
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Left= OfficeIMO.OfficeLength.Emus((long)Math.Round(current));
                current += shape.Width.ToEmus() + spacingEmus;
            }
        }

        private static void DistributeVerticalWithSpacing(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, bool center) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Top.ToEmus()).ToList();
            long totalHeight = ordered.Sum(s => s.Height.ToEmus());
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockHeight = totalHeight + totalSpacing;
            double start = bounds.Top.ToEmus();

            if (center && blockHeight < bounds.Height.ToEmus()) {
                start = bounds.Top.ToEmus() + (bounds.Height.ToEmus() - blockHeight) / 2d;
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Top= OfficeIMO.OfficeLength.Emus((long)Math.Round(current));
                current += shape.Height.ToEmus() + spacingEmus;
            }
        }

        private static void DistributeHorizontalWithSpacingAligned(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Left.ToEmus()).ToList();
            long totalWidth = ordered.Sum(s => s.Width.ToEmus());
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockWidth = totalWidth + totalSpacing;

            double start = bounds.Left.ToEmus();
            if (blockWidth <= bounds.Width.ToEmus()) {
                start = alignment switch {
                    PowerPointShapeAlignment.Left => bounds.Left.ToEmus(),
                    PowerPointShapeAlignment.Center => bounds.Left.ToEmus() + (bounds.Width.ToEmus() - blockWidth) / 2d,
                    PowerPointShapeAlignment.Right => bounds.Right.ToEmus() - blockWidth,
                    _ => throw new ArgumentOutOfRangeException(nameof(alignment))
                };
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Left= OfficeIMO.OfficeLength.Emus((long)Math.Round(current));
                current += shape.Width.ToEmus() + spacingEmus;
            }
        }

        private static void DistributeHorizontalWithSpacingAligned(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Left.ToEmus()).ToList();
            long totalWidth = ordered.Sum(s => s.Width.ToEmus());
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockWidth = totalWidth + totalSpacing;

            double start = bounds.Left.ToEmus();
            if (blockWidth <= bounds.Width.ToEmus()) {
                start = alignment switch {
                    PowerPointShapeAlignment.Left => bounds.Left.ToEmus(),
                    PowerPointShapeAlignment.Center => bounds.Left.ToEmus() + (bounds.Width.ToEmus() - blockWidth) / 2d,
                    PowerPointShapeAlignment.Right => bounds.Right.ToEmus() - blockWidth,
                    _ => throw new ArgumentOutOfRangeException(nameof(alignment))
                };
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Left= OfficeIMO.OfficeLength.Emus((long)Math.Round(current));
                current += shape.Width.ToEmus() + spacingEmus;
            }

            ApplyCrossAxisAlignment(ordered, PowerPointShapeDistribution.Horizontal, bounds, crossAxisAlignment);
        }

        private static void DistributeVerticalWithSpacingAligned(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Top.ToEmus()).ToList();
            long totalHeight = ordered.Sum(s => s.Height.ToEmus());
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockHeight = totalHeight + totalSpacing;

            double start = bounds.Top.ToEmus();
            if (blockHeight <= bounds.Height.ToEmus()) {
                start = alignment switch {
                    PowerPointShapeAlignment.Top => bounds.Top.ToEmus(),
                    PowerPointShapeAlignment.Middle => bounds.Top.ToEmus() + (bounds.Height.ToEmus() - blockHeight) / 2d,
                    PowerPointShapeAlignment.Bottom => bounds.Bottom.ToEmus() - blockHeight,
                    _ => throw new ArgumentOutOfRangeException(nameof(alignment))
                };
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Top= OfficeIMO.OfficeLength.Emus((long)Math.Round(current));
                current += shape.Height.ToEmus() + spacingEmus;
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
                throw new ArgumentOutOfRangeException("spacing");
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count < 2) {
                return;
            }

            ValidateLayoutBounds(bounds);
            ValidateDistributionLayout(distribution, alignment, crossAxisAlignment);
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
                    ? shapes[i].Width.ToEmus()
                    : shapes[i].Height.ToEmus();
            }

            long available = (distribution == PowerPointShapeDistribution.Horizontal ? bounds.Width.ToEmus() : bounds.Height.ToEmus()) - totalSize;
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
                    ? shapes[i].Width.ToEmus()
                    : shapes[i].Height.ToEmus();
            }

            long available = (distribution == PowerPointShapeDistribution.Horizontal ? bounds.Width.ToEmus() : bounds.Height.ToEmus())
                - spacingEmus * Math.Max(0, shapes.Count - 1);
            if (totalSize <= 0 || available >= totalSize) {
                return;
            }

            double scale = available <= 0 ? 0d : available / (double)totalSize;
            ScaleShapesForAxis(shapes, distribution == PowerPointShapeDistribution.Horizontal, scale, preserveAspect);
        }

        private static void DistributeVerticalWithSpacingAligned(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment, PowerPointShapeAlignment crossAxisAlignment) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Top.ToEmus()).ToList();
            long totalHeight = ordered.Sum(s => s.Height.ToEmus());
            long totalSpacing = spacingEmus * (ordered.Count - 1);
            double blockHeight = totalHeight + totalSpacing;

            double start = bounds.Top.ToEmus();
            if (blockHeight <= bounds.Height.ToEmus()) {
                start = alignment switch {
                    PowerPointShapeAlignment.Top => bounds.Top.ToEmus(),
                    PowerPointShapeAlignment.Middle => bounds.Top.ToEmus() + (bounds.Height.ToEmus() - blockHeight) / 2d,
                    PowerPointShapeAlignment.Bottom => bounds.Bottom.ToEmus() - blockHeight,
                    _ => throw new ArgumentOutOfRangeException(nameof(alignment))
                };
            }

            double current = start;
            foreach (PowerPointShape shape in ordered) {
                shape.Top= OfficeIMO.OfficeLength.Emus((long)Math.Round(current));
                current += shape.Height.ToEmus() + spacingEmus;
            }

            ApplyCrossAxisAlignment(ordered, PowerPointShapeDistribution.Vertical, bounds, crossAxisAlignment);
        }
    }
}
