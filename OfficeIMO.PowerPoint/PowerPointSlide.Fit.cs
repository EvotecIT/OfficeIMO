using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Scales and repositions shapes to fit within a target bounding box.
        /// </summary>
        public void FitShapesToBounds(IEnumerable<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            bool preserveAspect = false, bool center = false) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
            }

            PowerPointLayoutBox source = GetSelectionBounds(list);
            if (source.Width.ToEmus() == 0 || source.Height.ToEmus() == 0) {
                return;
            }

            double scaleX = bounds.Width.ToEmus() / (double)source.Width.ToEmus();
            double scaleY = bounds.Height.ToEmus() / (double)source.Height.ToEmus();

            if (preserveAspect) {
                double scale = Math.Min(scaleX, scaleY);
                scaleX = scale;
                scaleY = scale;
            }

            double scaledWidth = source.Width.ToEmus() * scaleX;
            double scaledHeight = source.Height.ToEmus() * scaleY;

            double baseLeft = bounds.Left.ToEmus();
            double baseTop = bounds.Top.ToEmus();

            if (center) {
                baseLeft += (bounds.Width.ToEmus() - scaledWidth) / 2d;
                baseTop += (bounds.Height.ToEmus() - scaledHeight) / 2d;
            }

            foreach (PowerPointShape shape in list) {
                double relativeLeft = shape.Left.ToEmus() - source.Left.ToEmus();
                double relativeTop = shape.Top.ToEmus() - source.Top.ToEmus();
                double newLeft = baseLeft + (relativeLeft * scaleX);
                double newTop = baseTop + (relativeTop * scaleY);

                shape.Left= OfficeIMO.OfficeLength.Emus((long)Math.Round(newLeft));
                shape.Top= OfficeIMO.OfficeLength.Emus((long)Math.Round(newTop));
                shape.Width= OfficeIMO.OfficeLength.Emus((long)Math.Round(shape.Width.ToEmus() * scaleX));
                shape.Height= OfficeIMO.OfficeLength.Emus((long)Math.Round(shape.Height.ToEmus() * scaleY));
            }
        }

        /// <summary>
        ///     Scales and repositions shapes to fit the slide bounds.
        /// </summary>
        public void FitShapesToSlide(IEnumerable<PowerPointShape> shapes, bool preserveAspect = false, bool center = false) {
            FitShapesToBounds(shapes, GetSlideBounds(), preserveAspect, center);
        }

        /// <summary>
        ///     Scales and repositions shapes to fit the slide content bounds with a margin (EMUs).
        /// </summary>
        public void FitShapesToSlideContent(IEnumerable<PowerPointShape> shapes, long marginEmus,
            bool preserveAspect = false, bool center = false) {
            FitShapesToBounds(shapes, GetSlideContentBounds(marginEmus), preserveAspect, center);
        }

        /// <summary>
        ///     Scales and repositions shapes to fit the slide content bounds with a margin (centimeters).
        /// </summary>
        public void FitShapesToSlideContentCm(IEnumerable<PowerPointShape> shapes, double marginCm,
            bool preserveAspect = false, bool center = false) {
            FitShapesToSlideContent(shapes, PowerPointUnits.FromCentimeters(marginCm), preserveAspect, center);
        }

        /// <summary>
        ///     Scales and repositions shapes to fit the slide content bounds with a margin (inches).
        /// </summary>
        public void FitShapesToSlideContentInches(IEnumerable<PowerPointShape> shapes, double marginInches,
            bool preserveAspect = false, bool center = false) {
            FitShapesToSlideContent(shapes, PowerPointUnits.FromInches(marginInches), preserveAspect, center);
        }

        /// <summary>
        ///     Scales and repositions shapes to fit the slide content bounds with a margin (points).
        /// </summary>
        public void FitShapesToSlideContentPoints(IEnumerable<PowerPointShape> shapes, double marginPoints,
            bool preserveAspect = false, bool center = false) {
            FitShapesToSlideContent(shapes, PowerPointUnits.FromPoints(marginPoints), preserveAspect, center);
        }

        private PowerPointLayoutBox GetSlideContentBounds(long marginEmus) {
            PresentationDocument? document = _slidePart.OpenXmlPackage as PresentationDocument;
            PresentationPart? presentationPart = document?.PresentationPart;
            if (presentationPart == null) {
                return GetSelectionBounds(ShapeList);
            }

            var slideSize = new PowerPointSlideSize(presentationPart);
            return slideSize.GetContentBox(marginEmus);
        }
    }
}
