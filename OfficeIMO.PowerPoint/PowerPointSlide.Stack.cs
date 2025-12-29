using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Stacks shapes within their selection bounds using a fixed spacing.
        /// </summary>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, long spacingEmus) {
            StackShapes(shapes, direction, GetSelectionBounds(shapes), spacingEmus,
                GetDefaultStackAlignment(direction), PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within their selection bounds using a fixed spacing and alignment.
        /// </summary>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, long spacingEmus,
            PowerPointShapeAlignment alignment) {
            StackShapes(shapes, direction, GetSelectionBounds(shapes), spacingEmus, alignment, PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within their selection bounds using a fixed spacing and justification.
        /// </summary>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, long spacingEmus,
            PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSelectionBounds(shapes), spacingEmus,
                GetDefaultStackAlignment(direction), justify);
        }

        /// <summary>
        ///     Stacks shapes within their selection bounds using a fixed spacing, alignment, and justification.
        /// </summary>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, long spacingEmus,
            PowerPointShapeAlignment alignment, PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSelectionBounds(shapes), spacingEmus, alignment, justify);
        }

        /// <summary>
        ///     Stacks shapes within a custom bounding box using a fixed spacing.
        /// </summary>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, PowerPointLayoutBox bounds,
            long spacingEmus) {
            StackShapes(shapes, direction, bounds, spacingEmus, GetDefaultStackAlignment(direction),
                PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within a custom bounding box using a fixed spacing and alignment.
        /// </summary>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment) {
            StackShapes(shapes, direction, bounds, spacingEmus, alignment, PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within a custom bounding box using a fixed spacing and justification.
        /// </summary>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, bounds, spacingEmus, GetDefaultStackAlignment(direction), justify);
        }

        /// <summary>
        ///     Stacks shapes within a custom bounding box using a fixed spacing, alignment, and justification.
        /// </summary>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment, PowerPointShapeStackJustify justify) {
            StackShapesInternal(shapes, direction, bounds, spacingEmus, alignment, justify,
                clampSpacingToBounds: false, scaleToFitBounds: false, preserveAspect: false);
        }

        /// <summary>
        ///     Stacks shapes within a custom bounding box using options.
        /// </summary>
        public void StackShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, PowerPointLayoutBox bounds,
            PowerPointShapeStackOptions options) {
            PowerPointShapeStackOptions resolvedOptions = options ?? new PowerPointShapeStackOptions();
            PowerPointShapeAlignment alignment = resolvedOptions.Alignment ?? GetDefaultStackAlignment(direction);
            StackShapesInternal(shapes, direction, bounds, resolvedOptions.SpacingEmus, alignment,
                resolvedOptions.Justify, resolvedOptions.ClampSpacingToBounds,
                resolvedOptions.ScaleToFitBounds, resolvedOptions.PreserveAspect);
        }

        /// <summary>
        ///     Stacks shapes within the slide bounds using a fixed spacing.
        /// </summary>
        public void StackShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, long spacingEmus) {
            StackShapes(shapes, direction, GetSlideBounds(), spacingEmus,
                GetDefaultStackAlignment(direction), PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within the slide bounds using a fixed spacing and alignment.
        /// </summary>
        public void StackShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, long spacingEmus,
            PowerPointShapeAlignment alignment) {
            StackShapes(shapes, direction, GetSlideBounds(), spacingEmus, alignment, PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within the slide bounds using a fixed spacing and justification.
        /// </summary>
        public void StackShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, long spacingEmus,
            PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSlideBounds(), spacingEmus, GetDefaultStackAlignment(direction), justify);
        }

        /// <summary>
        ///     Stacks shapes within the slide bounds using a fixed spacing, alignment, and justification.
        /// </summary>
        public void StackShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction, long spacingEmus,
            PowerPointShapeAlignment alignment, PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSlideBounds(), spacingEmus, alignment, justify);
        }

        /// <summary>
        ///     Stacks shapes within the slide bounds using options.
        /// </summary>
        public void StackShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            PowerPointShapeStackOptions options) {
            StackShapes(shapes, direction, GetSlideBounds(), options);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using a margin (EMUs).
        /// </summary>
        public void StackShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            long spacingEmus, long marginEmus) {
            StackShapes(shapes, direction, GetSlideContentBounds(marginEmus), spacingEmus,
                GetDefaultStackAlignment(direction), PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using a margin (EMUs) and alignment.
        /// </summary>
        public void StackShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            long spacingEmus, long marginEmus, PowerPointShapeAlignment alignment) {
            StackShapes(shapes, direction, GetSlideContentBounds(marginEmus), spacingEmus, alignment,
                PowerPointShapeStackJustify.Start);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using a margin (EMUs) and justification.
        /// </summary>
        public void StackShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            long spacingEmus, long marginEmus, PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSlideContentBounds(marginEmus), spacingEmus,
                GetDefaultStackAlignment(direction), justify);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using a margin (EMUs), alignment, and justification.
        /// </summary>
        public void StackShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            long spacingEmus, long marginEmus, PowerPointShapeAlignment alignment, PowerPointShapeStackJustify justify) {
            StackShapes(shapes, direction, GetSlideContentBounds(marginEmus), spacingEmus, alignment, justify);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using a margin (EMUs) and options.
        /// </summary>
        public void StackShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            long marginEmus, PowerPointShapeStackOptions options) {
            StackShapes(shapes, direction, GetSlideContentBounds(marginEmus), options);
        }

        /// <summary>
        ///     Stacks shapes using centimeters.
        /// </summary>
        public void StackShapesCm(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            double spacingCm) {
            StackShapes(shapes, direction, PowerPointUnits.FromCentimeters(spacingCm));
        }

        /// <summary>
        ///     Stacks shapes using inches.
        /// </summary>
        public void StackShapesInches(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            double spacingInches) {
            StackShapes(shapes, direction, PowerPointUnits.FromInches(spacingInches));
        }

        /// <summary>
        ///     Stacks shapes using points.
        /// </summary>
        public void StackShapesPoints(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            double spacingPoints) {
            StackShapes(shapes, direction, PowerPointUnits.FromPoints(spacingPoints));
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using centimeters.
        /// </summary>
        public void StackShapesToSlideContentCm(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            double spacingCm, double marginCm) {
            StackShapesToSlideContent(shapes, direction,
                PowerPointUnits.FromCentimeters(spacingCm),
                PowerPointUnits.FromCentimeters(marginCm));
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using inches.
        /// </summary>
        public void StackShapesToSlideContentInches(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            double spacingInches, double marginInches) {
            StackShapesToSlideContent(shapes, direction,
                PowerPointUnits.FromInches(spacingInches),
                PowerPointUnits.FromInches(marginInches));
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using points.
        /// </summary>
        public void StackShapesToSlideContentPoints(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            double spacingPoints, double marginPoints) {
            StackShapesToSlideContent(shapes, direction,
                PowerPointUnits.FromPoints(spacingPoints),
                PowerPointUnits.FromPoints(marginPoints));
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using centimeters and justification.
        /// </summary>
        public void StackShapesToSlideContentCm(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            double spacingCm, double marginCm, PowerPointShapeStackJustify justify) {
            StackShapesToSlideContent(shapes, direction,
                PowerPointUnits.FromCentimeters(spacingCm),
                PowerPointUnits.FromCentimeters(marginCm),
                justify);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using inches and justification.
        /// </summary>
        public void StackShapesToSlideContentInches(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            double spacingInches, double marginInches, PowerPointShapeStackJustify justify) {
            StackShapesToSlideContent(shapes, direction,
                PowerPointUnits.FromInches(spacingInches),
                PowerPointUnits.FromInches(marginInches),
                justify);
        }

        /// <summary>
        ///     Stacks shapes within the slide content bounds using points and justification.
        /// </summary>
        public void StackShapesToSlideContentPoints(IEnumerable<PowerPointShape> shapes, PowerPointShapeStackDirection direction,
            double spacingPoints, double marginPoints, PowerPointShapeStackJustify justify) {
            StackShapesToSlideContent(shapes, direction,
                PowerPointUnits.FromPoints(spacingPoints),
                PowerPointUnits.FromPoints(marginPoints),
                justify);
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
                throw new ArgumentOutOfRangeException(nameof(spacingEmus));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
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
                    ? shapes[i].Width
                    : shapes[i].Height;
            }

            long available = (direction == PowerPointShapeStackDirection.Horizontal ? bounds.Width : bounds.Height)
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
                    ? shapes[i].Width
                    : shapes[i].Height;
            }

            long available = (direction == PowerPointShapeStackDirection.Horizontal ? bounds.Width : bounds.Height) - totalSize;
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
                totalWidth += shapes[i].Width;
            }
            long totalSpacing = spacingEmus * Math.Max(0, shapes.Count - 1);
            double blockWidth = totalWidth + totalSpacing;
            double current = bounds.Left;

            if (blockWidth <= bounds.Width) {
                current = justify switch {
                    PowerPointShapeStackJustify.Start => bounds.Left,
                    PowerPointShapeStackJustify.Center => bounds.Left + (bounds.Width - blockWidth) / 2d,
                    PowerPointShapeStackJustify.End => bounds.Right - blockWidth,
                    _ => bounds.Left
                };
            }

            foreach (PowerPointShape shape in shapes) {
                shape.Left = (long)Math.Round(current);
                shape.Top = ResolveTop(bounds, shape, alignment);
                current += shape.Width + spacingEmus;
            }
        }

        private static void StackVertical(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long spacingEmus, PowerPointShapeAlignment alignment, PowerPointShapeStackJustify justify) {
            long totalHeight = 0;
            for (int i = 0; i < shapes.Count; i++) {
                totalHeight += shapes[i].Height;
            }
            long totalSpacing = spacingEmus * Math.Max(0, shapes.Count - 1);
            double blockHeight = totalHeight + totalSpacing;
            double current = bounds.Top;

            if (blockHeight <= bounds.Height) {
                current = justify switch {
                    PowerPointShapeStackJustify.Start => bounds.Top,
                    PowerPointShapeStackJustify.Center => bounds.Top + (bounds.Height - blockHeight) / 2d,
                    PowerPointShapeStackJustify.End => bounds.Bottom - blockHeight,
                    _ => bounds.Top
                };
            }

            foreach (PowerPointShape shape in shapes) {
                shape.Top = (long)Math.Round(current);
                shape.Left = ResolveLeft(bounds, shape, alignment);
                current += shape.Height + spacingEmus;
            }
        }

        private static long ResolveTop(PowerPointLayoutBox bounds, PowerPointShape shape, PowerPointShapeAlignment alignment) {
            return alignment switch {
                PowerPointShapeAlignment.Top => bounds.Top,
                PowerPointShapeAlignment.Middle => bounds.Top + (long)Math.Round((bounds.Height - shape.Height) / 2d),
                PowerPointShapeAlignment.Bottom => bounds.Bottom - shape.Height,
                _ => throw new ArgumentOutOfRangeException(nameof(alignment),
                    "Horizontal stacking supports Top, Middle, or Bottom alignment.")
            };
        }

        private static long ResolveLeft(PowerPointLayoutBox bounds, PowerPointShape shape, PowerPointShapeAlignment alignment) {
            return alignment switch {
                PowerPointShapeAlignment.Left => bounds.Left,
                PowerPointShapeAlignment.Center => bounds.Left + (long)Math.Round((bounds.Width - shape.Width) / 2d),
                PowerPointShapeAlignment.Right => bounds.Right - shape.Width,
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
                    shape.Width = (long)Math.Round(shape.Width * scale);
                    if (preserveAspect) {
                        shape.Height = (long)Math.Round(shape.Height * scale);
                    }
                } else {
                    shape.Height = (long)Math.Round(shape.Height * scale);
                    if (preserveAspect) {
                        shape.Width = (long)Math.Round(shape.Width * scale);
                    }
                }
            }
        }
    }
}
