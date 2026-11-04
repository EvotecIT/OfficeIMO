using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Aligns shapes within their selection bounds.
        /// </summary>
        public void AlignShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeAlignment alignment) {
            AlignShapes(shapes, alignment, GetSelectionBounds(shapes));
        }

        /// <summary>
        ///     Aligns shapes within a custom bounding box.
        /// </summary>
        public void AlignShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeAlignment alignment, PowerPointLayoutBox bounds) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
            }

            ValidateLayoutBounds(bounds);
            if (alignment < PowerPointShapeAlignment.Left || alignment > PowerPointShapeAlignment.Bottom) {
                throw new ArgumentOutOfRangeException(nameof(alignment));
            }

            foreach (PowerPointShape shape in list) {
                switch (alignment) {
                    case PowerPointShapeAlignment.Left:
                        shape.Left= OfficeIMO.OfficeLength.Emus(bounds.Left.ToEmus());
                        break;
                    case PowerPointShapeAlignment.Center:
                        shape.Left= OfficeIMO.OfficeLength.Emus(bounds.Left.ToEmus() + (long)Math.Round((bounds.Width.ToEmus() - shape.Width.ToEmus()) / 2d));
                        break;
                    case PowerPointShapeAlignment.Right:
                        shape.Left= OfficeIMO.OfficeLength.Emus(bounds.Right.ToEmus() - shape.Width.ToEmus());
                        break;
                    case PowerPointShapeAlignment.Top:
                        shape.Top= OfficeIMO.OfficeLength.Emus(bounds.Top.ToEmus());
                        break;
                    case PowerPointShapeAlignment.Middle:
                        shape.Top= OfficeIMO.OfficeLength.Emus(bounds.Top.ToEmus() + (long)Math.Round((bounds.Height.ToEmus() - shape.Height.ToEmus()) / 2d));
                        break;
                    case PowerPointShapeAlignment.Bottom:
                        shape.Top= OfficeIMO.OfficeLength.Emus(bounds.Bottom.ToEmus() - shape.Height.ToEmus());
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(alignment));
                }
            }
        }

        /// <summary>
        ///     Aligns shapes within the full slide bounds.
        /// </summary>
        public void AlignShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeAlignment alignment) {
            AlignShapes(shapes, alignment, GetSlideBounds());
        }

        /// <summary>
        ///     Aligns shapes within the slide content bounds using a margin.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="alignment">Alignment within the bounds on the applicable axis.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void AlignShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeAlignment alignment,
            OfficeLength margin) {
            AlignShapes(shapes, alignment, GetSlideContentBounds(margin));
        }

        /// <summary>
        ///     Distributes shapes evenly within their selection bounds.        
        /// </summary>
        public void DistributeShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution) {
            DistributeShapes(shapes, distribution, GetSelectionBounds(shapes));
        }

        /// <summary>
        ///     Distributes shapes evenly within their selection bounds and aligns them on the cross axis.
        /// </summary>
        public void DistributeShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapes(shapes, distribution, GetSelectionBounds(shapes), crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes evenly within a custom bounding box.
        /// </summary>
        public void DistributeShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution, PowerPointLayoutBox bounds) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count < 2) {
                return;
            }

            ValidateLayoutBounds(bounds);
            ValidateDistributionLayout(distribution);
            switch (distribution) {
                case PowerPointShapeDistribution.Horizontal:
                    DistributeHorizontal(list, bounds);
                    break;
                case PowerPointShapeDistribution.Vertical:
                    DistributeVertical(list, bounds);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(distribution));
            }
        }

        /// <summary>
        ///     Distributes shapes evenly within a custom bounding box and aligns them on the cross axis.
        /// </summary>
        public void DistributeShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, PowerPointShapeAlignment crossAxisAlignment) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
            }

            ValidateLayoutBounds(bounds);
            ValidateDistributionLayout(distribution, crossAxisAlignment: crossAxisAlignment);
            DistributeShapes(list, distribution, bounds);
            ApplyCrossAxisAlignment(list, distribution, bounds, crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes evenly within the full slide bounds.
        /// </summary>
        public void DistributeShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution) {
            DistributeShapes(shapes, distribution, GetSlideBounds());
        }

        /// <summary>
        ///     Distributes shapes evenly within the full slide bounds and aligns them on the cross axis.
        /// </summary>
        public void DistributeShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapes(shapes, distribution, GetSlideBounds(), crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes evenly within the slide content bounds using a margin and aligns them on the cross axis.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="crossAxisAlignment">Alignment on the axis perpendicular to distribution.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            OfficeLength margin, PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapes(shapes, distribution, GetSlideContentBounds(margin), crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes evenly within the slide content bounds using a margin.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            OfficeLength margin) {
            DistributeShapes(shapes, distribution, GetSlideContentBounds(margin));
        }

        // Validate both axes before distribution or scaling changes a shape.
        private static void ValidateDistributionLayout(PowerPointShapeDistribution distribution,
            PowerPointShapeAlignment? alignment = null, PowerPointShapeAlignment? crossAxisAlignment = null) {
            if (distribution != PowerPointShapeDistribution.Horizontal && distribution != PowerPointShapeDistribution.Vertical) {
                throw new ArgumentOutOfRangeException(nameof(distribution));
            }
            bool horizontal = distribution == PowerPointShapeDistribution.Horizontal;
            if (alignment.HasValue && !IsLayoutAlignment(alignment.Value, horizontal)) {
                throw new ArgumentOutOfRangeException(nameof(alignment));
            }
            if (crossAxisAlignment.HasValue && !IsLayoutAlignment(crossAxisAlignment.Value, !horizontal)) {
                throw new ArgumentOutOfRangeException(nameof(crossAxisAlignment));
            }
        }

        private static bool IsLayoutAlignment(PowerPointShapeAlignment alignment, bool horizontal) {
            return horizontal
                ? alignment == PowerPointShapeAlignment.Left || alignment == PowerPointShapeAlignment.Center || alignment == PowerPointShapeAlignment.Right
                : alignment == PowerPointShapeAlignment.Top || alignment == PowerPointShapeAlignment.Middle || alignment == PowerPointShapeAlignment.Bottom;
        }

        private static void ValidateLayoutBounds(PowerPointLayoutBox bounds) {
            if (bounds.Width < OfficeLength.Zero || bounds.Height < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(bounds), "Layout dimensions cannot be negative.");
            }
            // Evaluate checked edges before any shape is repositioned.
            _ = bounds.Right;
            _ = bounds.Bottom;
        }

        private static List<PowerPointShape> NormalizeShapes(IEnumerable<PowerPointShape> shapes) {
            return shapes.Where(s => s != null).ToList();
        }

        private static PowerPointLayoutBox GetSelectionBounds(IEnumerable<PowerPointShape> shapes) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0));
            }

            long left = list.Min(s => s.Left.ToEmus());
            long top = list.Min(s => s.Top.ToEmus());
            long right = list.Max(s => s.Left.ToEmus() + s.Width.ToEmus());
            long bottom = list.Max(s => s.Top.ToEmus() + s.Height.ToEmus());
            return new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(left),OfficeIMO.OfficeLength.Emus(top),OfficeIMO.OfficeLength.Emus(right - left),OfficeIMO.OfficeLength.Emus(bottom - top));
        }

        private PowerPointLayoutBox GetSlideBounds() {
            PresentationDocument? document = _slidePart.OpenXmlPackage as PresentationDocument;
            PresentationPart? presentationPart = document?.PresentationPart;
            if (presentationPart == null) {
                return GetSelectionBounds(ShapeList);
            }

            var slideSize = new PowerPointSlideSize(presentationPart);
            return new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(slideSize.Width.ToEmus()),OfficeIMO.OfficeLength.Emus(slideSize.Height.ToEmus()));
        }

        private static void DistributeHorizontal(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Left.ToEmus()).ToList();
            long totalWidth = ordered.Sum(s => s.Width.ToEmus());
            double available = bounds.Width.ToEmus() - totalWidth;
            double gap = ordered.Count > 1 ? available / (ordered.Count - 1) : 0d;

            double current = bounds.Left.ToEmus();
            for (int i = 0; i < ordered.Count; i++) {
                PowerPointShape shape = ordered[i];
                if (i == ordered.Count - 1 && ordered.Count > 1) {
                    shape.Left= OfficeIMO.OfficeLength.Emus(bounds.Right.ToEmus() - shape.Width.ToEmus());
                } else {
                    shape.Left= OfficeIMO.OfficeLength.Emus((long)Math.Round(current));
                }
                current += shape.Width.ToEmus() + gap;
            }
        }

        private static void DistributeVertical(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Top.ToEmus()).ToList();
            long totalHeight = ordered.Sum(s => s.Height.ToEmus());
            double available = bounds.Height.ToEmus() - totalHeight;
            double gap = ordered.Count > 1 ? available / (ordered.Count - 1) : 0d;

            double current = bounds.Top.ToEmus();
            for (int i = 0; i < ordered.Count; i++) {
                PowerPointShape shape = ordered[i];
                if (i == ordered.Count - 1 && ordered.Count > 1) {
                    shape.Top= OfficeIMO.OfficeLength.Emus(bounds.Bottom.ToEmus() - shape.Height.ToEmus());
                } else {
                    shape.Top= OfficeIMO.OfficeLength.Emus((long)Math.Round(current));
                }
                current += shape.Height.ToEmus() + gap;
            }
        }

        private static void ApplyCrossAxisAlignment(IReadOnlyList<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, PowerPointShapeAlignment alignment) {
            switch (distribution) {
                case PowerPointShapeDistribution.Horizontal:
                    foreach (PowerPointShape shape in shapes) {
                        shape.Top= OfficeIMO.OfficeLength.Emus(ResolveCrossAxisTop(bounds, shape, alignment));
                    }
                    break;
                case PowerPointShapeDistribution.Vertical:
                    foreach (PowerPointShape shape in shapes) {
                        shape.Left= OfficeIMO.OfficeLength.Emus(ResolveCrossAxisLeft(bounds, shape, alignment));
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(distribution));
            }
        }

        private static long ResolveCrossAxisTop(PowerPointLayoutBox bounds, PowerPointShape shape,
            PowerPointShapeAlignment alignment) {
            return alignment switch {
                PowerPointShapeAlignment.Top => bounds.Top.ToEmus(),
                PowerPointShapeAlignment.Middle => bounds.Top.ToEmus() + (long)Math.Round((bounds.Height.ToEmus() - shape.Height.ToEmus()) / 2d),
                PowerPointShapeAlignment.Bottom => bounds.Bottom.ToEmus() - shape.Height.ToEmus(),
                _ => throw new ArgumentOutOfRangeException(nameof(alignment),
                    "Horizontal distribution supports Top, Middle, or Bottom alignment.")
            };
        }

        private static long ResolveCrossAxisLeft(PowerPointLayoutBox bounds, PowerPointShape shape,
            PowerPointShapeAlignment alignment) {
            return alignment switch {
                PowerPointShapeAlignment.Left => bounds.Left.ToEmus(),
                PowerPointShapeAlignment.Center => bounds.Left.ToEmus() + (long)Math.Round((bounds.Width.ToEmus() - shape.Width.ToEmus()) / 2d),
                PowerPointShapeAlignment.Right => bounds.Right.ToEmus() - shape.Width.ToEmus(),
                _ => throw new ArgumentOutOfRangeException(nameof(alignment),
                    "Vertical distribution supports Left, Center, or Right alignment.")
            };
        }
    }
}
