using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Stacks shapes within their selection bounds using a fixed spacing.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, OfficeLength spacing) {
            StackShapes(shapes, direction, GetSelectionBounds(shapes), spacing,
                GetDefaultStackAlignment(direction), PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within their selection bounds using a fixed spacing and alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, OfficeLength spacing,
            PowerPointShapeAlignment alignment) {
            StackShapes(shapes, direction, GetSelectionBounds(shapes), spacing, alignment, PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within their selection bounds using a fixed spacing and justification.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="justify">Position of the stack along its main axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, OfficeLength spacing,
            PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSelectionBounds(shapes), spacing,
                GetDefaultStackAlignment(direction), justify);
        }

        /// <summary>
        ///     Stacks shapes within their selection bounds using a fixed spacing, alignment, and justification.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <param name="justify">Position of the stack along its main axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, OfficeLength spacing,
            PowerPointShapeAlignment alignment, PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSelectionBounds(shapes), spacing, alignment, justify);
        }

        /// <summary>
        ///     Stacks shapes within a custom bounding box using a fixed spacing.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, PowerPointLayoutBox bounds,
            OfficeLength spacing) {
            StackShapes(shapes, direction, bounds, spacing, GetDefaultStackAlignment(direction),
                PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within a custom bounding box using a fixed spacing and alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, PowerPointLayoutBox bounds,
            OfficeLength spacing, PowerPointShapeAlignment alignment) {
            StackShapes(shapes, direction, bounds, spacing, alignment, PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within a custom bounding box using a fixed spacing and justification.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="justify">Position of the stack along its main axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, PowerPointLayoutBox bounds,
            OfficeLength spacing, PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, bounds, spacing, GetDefaultStackAlignment(direction), justify);
        }

        /// <summary>
        ///     Stacks shapes within a custom bounding box using a fixed spacing, alignment, and justification.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <param name="justify">Position of the stack along its main axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, PowerPointLayoutBox bounds,
            OfficeLength spacing, PowerPointShapeAlignment alignment, PowerPointShapeStackJustify justify) {
            StackShapesInternal(shapes, direction, bounds, spacing.ToEmus(), alignment, justify,
                clampSpacingToBounds: false, scaleToFitBounds: false, preserveAspect: false);
        }

        /// <summary>
        ///     Stacks shapes within a custom bounding box using options.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, PowerPointLayoutBox bounds,
            PowerPointShapeStackOptions? options) {
            PowerPointShapeStackOptions resolvedOptions = options ?? new PowerPointShapeStackOptions();
            PowerPointShapeAlignment alignment = resolvedOptions.Alignment ?? GetDefaultStackAlignment(direction);
            StackShapesInternal(shapes, direction, bounds, resolvedOptions.Spacing.ToEmus(), alignment,
                resolvedOptions.Justify, resolvedOptions.ClampSpacingToBounds,
                resolvedOptions.ScaleToFitBounds, resolvedOptions.PreserveAspect);
        }

        /// <summary>
        ///     Stacks shapes within the slide bounds using a fixed spacing.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, OfficeLength spacing) {
            StackShapes(shapes, direction, GetSlideBounds(), spacing,
                GetDefaultStackAlignment(direction), PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within the slide bounds using a fixed spacing and alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, OfficeLength spacing,
            PowerPointShapeAlignment alignment) {
            StackShapes(shapes, direction, GetSlideBounds(), spacing, alignment, PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within the slide bounds using a fixed spacing and justification.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="justify">Position of the stack along its main axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, OfficeLength spacing,
            PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSlideBounds(), spacing, GetDefaultStackAlignment(direction), justify);
        }

        /// <summary>
        ///     Stacks shapes within the slide bounds using a fixed spacing, alignment, and justification.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <param name="justify">Position of the stack along its main axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, OfficeLength spacing,
            PowerPointShapeAlignment alignment, PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSlideBounds(), spacing, alignment, justify);
        }

        /// <summary>
        ///     Stacks shapes within the slide bounds using options.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        public void StackShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            PowerPointShapeStackOptions? options) {
            StackShapes(shapes, direction, GetSlideBounds(), options);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using a margin.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            OfficeLength spacing, OfficeLength margin) {
            StackShapes(shapes, direction, GetSlideContentBounds(margin), spacing,
                GetDefaultStackAlignment(direction), PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using a margin and alignment.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            OfficeLength spacing, OfficeLength margin, PowerPointShapeAlignment alignment) {
            StackShapes(shapes, direction, GetSlideContentBounds(margin), spacing, alignment,
                PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using a margin and justification.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="justify">Position of the stack along its main axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            OfficeLength spacing, OfficeLength margin, PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSlideContentBounds(margin), spacing,
                GetDefaultStackAlignment(direction), justify);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using a margin, alignment, and justification.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <param name="justify">Position of the stack along its main axis.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            OfficeLength spacing, OfficeLength margin, PowerPointShapeAlignment alignment, PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSlideContentBounds(margin), spacing, alignment, justify);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using a margin and options.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            OfficeLength margin, PowerPointShapeStackOptions? options) {
            StackShapes(shapes, direction, GetSlideContentBounds(margin), options);
        }

        private static PowerPointShapeAlignment GetDefaultStackAlignment(PowerPointShapeStackDirection direction) {
            return direction == PowerPointShapeStackDirection.Horizontal
                ? PowerPointShapeAlignment.Top
                : PowerPointShapeAlignment.Left;
        }

        private void StackShapesInternal(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            PowerPointLayoutBox bounds, long spacingEmus, PowerPointShapeAlignment alignment,
            PowerPointShapeStackJustify justify, bool clampSpacingToBounds, bool scaleToFitBounds, bool preserveAspect) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }
            if (spacingEmus < 0) {
                throw new ArgumentOutOfRangeException("spacing");
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
            }

            ValidateLayoutBounds(bounds);
            if (direction != PowerPointShapeStackDirection.Horizontal && direction != PowerPointShapeStackDirection.Vertical) {
                throw new ArgumentOutOfRangeException(nameof(direction));
            }
            if (!IsLayoutAlignment(alignment, direction == PowerPointShapeStackDirection.Vertical)) {
                throw new ArgumentOutOfRangeException(nameof(alignment));
            }
            if (justify != PowerPointShapeStackJustify.Start && justify != PowerPointShapeStackJustify.Center && justify != PowerPointShapeStackJustify.End) {
                throw new ArgumentOutOfRangeException(nameof(justify));
            }
            long effectiveSpacing = ResolveSpacingForBounds(list, direction, bounds, spacingEmus, clampSpacingToBounds);
            if (scaleToFitBounds) {
                ApplyScaleToFitBounds(list, direction, bounds, effectiveSpacing, preserveAspect);
            }

            switch (direction) {
                case PowerPointShapeStackDirection.Horizontal:
                    StackHorizontal(list, bounds, effectiveSpacing, alignment, justify);
                    break;
                case PowerPointShapeStackDirection.Vertical:
                    StackVertical(list, bounds, effectiveSpacing, alignment, justify);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(direction));
            }
        }

        private static void ApplyScaleToFitBounds(IReadOnlyList<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            PowerPointLayoutBox bounds, long spacingEmus, bool preserveAspect) {
            if (shapes.Count == 0) {
                return;
            }

            long totalSize = 0;
            for (int i = 0; i < shapes.Count; i++) {
                totalSize += direction == PowerPointShapeStackDirection.Horizontal
                    ? shapes[i].Width.ToEmus()
                    : shapes[i].Height.ToEmus();
            }

            long available = (direction == PowerPointShapeStackDirection.Horizontal ? bounds.Width.ToEmus() : bounds.Height.ToEmus())
                - spacingEmus * Math.Max(0, shapes.Count - 1);
            if (totalSize <= 0 || available >= totalSize) {
                return;
            }

            double scale = available <= 0 ? 0d : available / (double)totalSize;
            ScaleShapesForAxis(shapes, direction == PowerPointShapeStackDirection.Horizontal, scale, preserveAspect);
        }

        private static long ResolveSpacingForBounds(IReadOnlyList<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            PowerPointLayoutBox bounds, long spacingEmus, bool clampSpacingToBounds) {
            if (!clampSpacingToBounds || shapes.Count < 2) {
                return spacingEmus;
            }

            long totalSize = 0;
            for (int i = 0; i < shapes.Count; i++) {
                totalSize += direction == PowerPointShapeStackDirection.Horizontal
                    ? shapes[i].Width.ToEmus()
                    : shapes[i].Height.ToEmus();
            }

            long available = (direction == PowerPointShapeStackDirection.Horizontal ? bounds.Width.ToEmus() : bounds.Height.ToEmus()) - totalSize;
            if (available <= 0) {
                return 0;
            }

            long maxSpacing = (long)Math.Floor(available / (double)(shapes.Count - 1));
            if (maxSpacing < 0) {
                return 0;
            }

            return Math.Min(spacingEmus, maxSpacing);
        }

        private static void StackHorizontal(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment, PowerPointShapeStackJustify justify) {
            long totalWidth = 0;
            for (int i = 0; i < shapes.Count; i++) {
                totalWidth += shapes[i].Width.ToEmus();
            }
            long totalSpacing = spacingEmus * Math.Max(0, shapes.Count - 1);
            double blockWidth = totalWidth + totalSpacing;
            double current = bounds.Left.ToEmus();

            if (blockWidth <= bounds.Width.ToEmus()) {
                current = justify switch {
                    PowerPointShapeStackJustify.Start => bounds.Left.ToEmus(),
                    PowerPointShapeStackJustify.Center => bounds.Left.ToEmus() + (bounds.Width.ToEmus() - blockWidth) / 2d,
                    PowerPointShapeStackJustify.End => bounds.Right.ToEmus() - blockWidth,
                    _ => bounds.Left.ToEmus()
                };
            }

            foreach (PowerPointShape shape in shapes) {
                shape.Left= OfficeIMO.OfficeLength.Emus((long)Math.Round(current));
                shape.Top= OfficeIMO.OfficeLength.Emus(ResolveTop(bounds, shape, alignment));
                current += shape.Width.ToEmus() + spacingEmus;
            }
        }

        private static void StackVertical(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment, PowerPointShapeStackJustify justify) {
            long totalHeight = 0;
            for (int i = 0; i < shapes.Count; i++) {
                totalHeight += shapes[i].Height.ToEmus();
            }
            long totalSpacing = spacingEmus * Math.Max(0, shapes.Count - 1);
            double blockHeight = totalHeight + totalSpacing;
            double current = bounds.Top.ToEmus();

            if (blockHeight <= bounds.Height.ToEmus()) {
                current = justify switch {
                    PowerPointShapeStackJustify.Start => bounds.Top.ToEmus(),
                    PowerPointShapeStackJustify.Center => bounds.Top.ToEmus() + (bounds.Height.ToEmus() - blockHeight) / 2d,
                    PowerPointShapeStackJustify.End => bounds.Bottom.ToEmus() - blockHeight,
                    _ => bounds.Top.ToEmus()
                };
            }

            foreach (PowerPointShape shape in shapes) {
                shape.Top= OfficeIMO.OfficeLength.Emus((long)Math.Round(current));
                shape.Left= OfficeIMO.OfficeLength.Emus(ResolveLeft(bounds, shape, alignment));
                current += shape.Height.ToEmus() + spacingEmus;
            }
        }

        private static long ResolveTop(PowerPointLayoutBox bounds, PowerPointShape shape, PowerPointShapeAlignment alignment) {
            return alignment switch {
                PowerPointShapeAlignment.Top => bounds.Top.ToEmus(),
                PowerPointShapeAlignment.Middle => bounds.Top.ToEmus() + (long)Math.Round((bounds.Height.ToEmus() - shape.Height.ToEmus()) / 2d),
                PowerPointShapeAlignment.Bottom => bounds.Bottom.ToEmus() - shape.Height.ToEmus(),
                _ => throw new ArgumentOutOfRangeException(nameof(alignment),
                    "Horizontal stacking supports Top, Middle, or Bottom alignment.")
            };
        }

        private static long ResolveLeft(PowerPointLayoutBox bounds, PowerPointShape shape, PowerPointShapeAlignment alignment) {
            return alignment switch {
                PowerPointShapeAlignment.Left => bounds.Left.ToEmus(),
                PowerPointShapeAlignment.Center => bounds.Left.ToEmus() + (long)Math.Round((bounds.Width.ToEmus() - shape.Width.ToEmus()) / 2d),
                PowerPointShapeAlignment.Right => bounds.Right.ToEmus() - shape.Width.ToEmus(),
                _ => throw new ArgumentOutOfRangeException(nameof(alignment),
                    "Vertical stacking supports Left, Center, or Right alignment.")
            };
        }

        private static void ScaleShapesForAxis(IReadOnlyList<PowerPointShape> shapes, bool horizontal, double scale,
            bool preserveAspect) {
            if (scale >= 1d) {
                return;
            }

            for (int i = 0; i < shapes.Count; i++) {
                PowerPointShape shape = shapes[i];
                if (horizontal) {
                    shape.Width= OfficeIMO.OfficeLength.Emus((long)Math.Round(shape.Width.ToEmus() * scale));
                    if (preserveAspect) {
                        shape.Height= OfficeIMO.OfficeLength.Emus((long)Math.Round(shape.Height.ToEmus() * scale));
                    }
                } else {
                    shape.Height= OfficeIMO.OfficeLength.Emus((long)Math.Round(shape.Height.ToEmus() * scale));
                    if (preserveAspect) {
                        shape.Width= OfficeIMO.OfficeLength.Emus((long)Math.Round(shape.Width.ToEmus() * scale));
                    }
                }
            }
        }
    }
}
