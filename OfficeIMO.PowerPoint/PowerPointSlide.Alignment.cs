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

            foreach (PowerPointShape shape in list) {
                switch (alignment) {
                    case PowerPointShapeAlignment.Left:
                        shape.Left = bounds.Left;
                        break;
                    case PowerPointShapeAlignment.Center:
                        shape.Left = bounds.Left + (long)Math.Round((bounds.Width - shape.Width) / 2d);
                        break;
                    case PowerPointShapeAlignment.Right:
                        shape.Left = bounds.Right - shape.Width;
                        break;
                    case PowerPointShapeAlignment.Top:
                        shape.Top = bounds.Top;
                        break;
                    case PowerPointShapeAlignment.Middle:
                        shape.Top = bounds.Top + (long)Math.Round((bounds.Height - shape.Height) / 2d);
                        break;
                    case PowerPointShapeAlignment.Bottom:
                        shape.Top = bounds.Bottom - shape.Height;
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
        ///     Aligns shapes within the slide content bounds using a margin (EMUs).
        /// </summary>
        public void AlignShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeAlignment alignment,
            long marginEmus) {
            AlignShapes(shapes, alignment, GetSlideContentBounds(marginEmus));
        }

        /// <summary>
        ///     Aligns shapes within the slide content bounds using a margin (centimeters).
        /// </summary>
        public void AlignShapesToSlideContentCm(IEnumerable<PowerPointShape> shapes, PowerPointShapeAlignment alignment,
            double marginCm) {
            AlignShapesToSlideContent(shapes, alignment, PowerPointUnits.FromCentimeters(marginCm));
        }

        /// <summary>
        ///     Aligns shapes within the slide content bounds using a margin (inches).
        /// </summary>
        public void AlignShapesToSlideContentInches(IEnumerable<PowerPointShape> shapes, PowerPointShapeAlignment alignment,
            double marginInches) {
            AlignShapesToSlideContent(shapes, alignment, PowerPointUnits.FromInches(marginInches));
        }

        /// <summary>
        ///     Aligns shapes within the slide content bounds using a margin (points).
        /// </summary>
        public void AlignShapesToSlideContentPoints(IEnumerable<PowerPointShape> shapes, PowerPointShapeAlignment alignment,
            double marginPoints) {
            AlignShapesToSlideContent(shapes, alignment, PowerPointUnits.FromPoints(marginPoints));
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
        ///     Distributes shapes evenly within the slide content bounds using a margin (EMUs) and aligns them on the cross axis.
        /// </summary>
        public void DistributeShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            long marginEmus, PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapes(shapes, distribution, GetSlideContentBounds(marginEmus), crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes evenly within the slide content bounds using a margin (centimeters) and aligns them on the cross axis.
        /// </summary>
        public void DistributeShapesToSlideContentCm(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            double marginCm, PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapesToSlideContent(shapes, distribution, PowerPointUnits.FromCentimeters(marginCm), crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes evenly within the slide content bounds using a margin (inches) and aligns them on the cross axis.
        /// </summary>
        public void DistributeShapesToSlideContentInches(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            double marginInches, PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapesToSlideContent(shapes, distribution, PowerPointUnits.FromInches(marginInches), crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes evenly within the slide content bounds using a margin (points) and aligns them on the cross axis.
        /// </summary>
        public void DistributeShapesToSlideContentPoints(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            double marginPoints, PowerPointShapeAlignment crossAxisAlignment) {
            DistributeShapesToSlideContent(shapes, distribution, PowerPointUnits.FromPoints(marginPoints), crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes shapes evenly within the slide content bounds using a margin (EMUs).
        /// </summary>
        public void DistributeShapesToSlideContent(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            long marginEmus) {
            DistributeShapes(shapes, distribution, GetSlideContentBounds(marginEmus));
        }

        /// <summary>
        ///     Distributes shapes evenly within the slide content bounds using a margin (centimeters).
        /// </summary>
        public void DistributeShapesToSlideContentCm(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            double marginCm) {
            DistributeShapesToSlideContent(shapes, distribution, PowerPointUnits.FromCentimeters(marginCm));
        }

        /// <summary>
        ///     Distributes shapes evenly within the slide content bounds using a margin (inches).
        /// </summary>
        public void DistributeShapesToSlideContentInches(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            double marginInches) {
            DistributeShapesToSlideContent(shapes, distribution, PowerPointUnits.FromInches(marginInches));
        }

        /// <summary>
        ///     Distributes shapes evenly within the slide content bounds using a margin (points).
        /// </summary>
        public void DistributeShapesToSlideContentPoints(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            double marginPoints) {
            DistributeShapesToSlideContent(shapes, distribution, PowerPointUnits.FromPoints(marginPoints));
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
                return new PowerPointLayoutBox(0, 0, 0, 0);
            }

            long left = list.Min(s => s.Left);
            long top = list.Min(s => s.Top);
            long right = list.Max(s => s.Left + s.Width);
            long bottom = list.Max(s => s.Top + s.Height);
            return new PowerPointLayoutBox(left, top, right - left, bottom - top);
        }

        private PowerPointLayoutBox GetSlideBounds() {
            PresentationDocument? document = _slidePart.OpenXmlPackage as PresentationDocument;
            PresentationPart? presentationPart = document?.PresentationPart;
            if (presentationPart == null) {
                return GetSelectionBounds(_shapes);
            }

            var slideSize = new PowerPointSlideSize(presentationPart);
            return new PowerPointLayoutBox(0, 0, slideSize.WidthEmus, slideSize.HeightEmus);
        }

        private static void DistributeHorizontal(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Left).ToList();
            long totalWidth = ordered.Sum(s => s.Width);
            double available = bounds.Width - totalWidth;
            double gap = ordered.Count > 1 ? available / (ordered.Count - 1) : 0d;

            double current = bounds.Left;
            for (int i = 0; i < ordered.Count; i++) {
                PowerPointShape shape = ordered[i];
                if (i == ordered.Count - 1 && ordered.Count > 1) {
                    shape.Left = bounds.Right - shape.Width;
                } else {
                    shape.Left = (long)Math.Round(current);
                }
                current += shape.Width + gap;
            }
        }

        private static void DistributeVertical(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Top).ToList();
            long totalHeight = ordered.Sum(s => s.Height);
            double available = bounds.Height - totalHeight;
            double gap = ordered.Count > 1 ? available / (ordered.Count - 1) : 0d;

            double current = bounds.Top;
            for (int i = 0; i < ordered.Count; i++) {
                PowerPointShape shape = ordered[i];
                if (i == ordered.Count - 1 && ordered.Count > 1) {
                    shape.Top = bounds.Bottom - shape.Height;
                } else {
                    shape.Top = (long)Math.Round(current);
                }
                current += shape.Height + gap;
            }
        }

        private static void ApplyCrossAxisAlignment(IReadOnlyList<PowerPointShape> shapes, PowerPointShapeDistribution distribution,
            PowerPointLayoutBox bounds, PowerPointShapeAlignment alignment) {
            switch (distribution) {
                case PowerPointShapeDistribution.Horizontal:
                    foreach (PowerPointShape shape in shapes) {
                        shape.Top = ResolveCrossAxisTop(bounds, shape, alignment);
                    }
                    break;
                case PowerPointShapeDistribution.Vertical:
                    foreach (PowerPointShape shape in shapes) {
                        shape.Left = ResolveCrossAxisLeft(bounds, shape, alignment);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(distribution));
            }
        }

        private static long ResolveCrossAxisTop(PowerPointLayoutBox bounds, PowerPointShape shape,
            PowerPointShapeAlignment alignment) {
            return alignment switch {
                PowerPointShapeAlignment.Top => bounds.Top,
                PowerPointShapeAlignment.Middle => bounds.Top + (long)Math.Round((bounds.Height - shape.Height) / 2d),
                PowerPointShapeAlignment.Bottom => bounds.Bottom - shape.Height,
                _ => throw new ArgumentOutOfRangeException(nameof(alignment),
                    "Horizontal distribution supports Top, Middle, or Bottom alignment.")
            };
        }

        private static long ResolveCrossAxisLeft(PowerPointLayoutBox bounds, PowerPointShape shape,
            PowerPointShapeAlignment alignment) {
            return alignment switch {
                PowerPointShapeAlignment.Left => bounds.Left,
                PowerPointShapeAlignment.Center => bounds.Left + (long)Math.Round((bounds.Width - shape.Width) / 2d),
                PowerPointShapeAlignment.Right => bounds.Right - shape.Width,
                _ => throw new ArgumentOutOfRangeException(nameof(alignment),
                    "Vertical distribution supports Left, Center, or Right alignment.")
            };
        }
    }
}
