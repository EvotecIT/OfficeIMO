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
        ///     Distributes shapes evenly within their selection bounds.
        /// </summary>
        public void DistributeShapes(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution) {
            DistributeShapes(shapes, distribution, GetSelectionBounds(shapes));
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
        ///     Distributes shapes evenly within the full slide bounds.
        /// </summary>
        public void DistributeShapesToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeDistribution distribution) {
            DistributeShapes(shapes, distribution, GetSlideBounds());
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
            if (gap < 0) {
                gap = 0;
            }

            double current = bounds.Left;
            foreach (PowerPointShape shape in ordered) {
                shape.Left = (long)Math.Round(current);
                current += shape.Width + gap;
            }
        }

        private static void DistributeVertical(IReadOnlyList<PowerPointShape> shapes, PowerPointLayoutBox bounds) {
            List<PowerPointShape> ordered = shapes.OrderBy(s => s.Top).ToList();
            long totalHeight = ordered.Sum(s => s.Height);
            double available = bounds.Height - totalHeight;
            double gap = ordered.Count > 1 ? available / (ordered.Count - 1) : 0d;
            if (gap < 0) {
                gap = 0;
            }

            double current = bounds.Top;
            foreach (PowerPointShape shape in ordered) {
                shape.Top = (long)Math.Round(current);
                current += shape.Height + gap;
            }
        }
    }
}
