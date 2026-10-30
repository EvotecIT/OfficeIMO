using System;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointLayoutEngine {
        /// <summary>Layout owns the frame geometry; the caller's native figure and its parts remain intact.</summary>
        private static void ArrangeNativeFigure(PowerPointSlide slide, PowerPointShape figure, PowerPointLayoutRect box) {
            slide.AdoptLayoutShape(figure);
            figure.Left = PowerPointUnits.FromPoints(box.X);
            figure.Top = PowerPointUnits.FromPoints(box.Y);
            figure.Width = PowerPointUnits.FromPoints(box.Width);
            figure.Height = PowerPointUnits.FromPoints(box.Height);
        }

        private static PowerPointLayoutChart CopyRepeatedNativeChart(PowerPointLayoutChart chart) =>
            new(chart.Kind, chart.Data) {
                Title = chart.Title, AlternativeText = chart.AlternativeText,
                NativeChart = (PowerPointChart)chart.NativeChart!.OwnerSlide!.DuplicateShape(chart.NativeChart)
            };

        private static PowerPointLayoutImage CopyRepeatedNativeImage(PowerPointLayoutImage image) {
            var copy = new PowerPointLayoutImage(image.Path, image.AlternativeText) {
                FocalX = image.FocalX, FocalY = image.FocalY, Contain = image.Contain, Caption = image.Caption,
                NativePicture = (PowerPointPicture)image.NativePicture!.OwnerSlide!.DuplicateShape(image.NativePicture)
            };
            copy.Annotations.AddRange(image.Annotations);
            return copy;
        }

        private static PowerPointPicture RenderNativeImage(PowerPointLayoutContext context, PowerPointLayoutWriter writer,
            PowerPointLayoutImage image, PowerPointLayoutRect box, bool cover) {
            PowerPointPicture picture = image.NativePicture!;
            bool hasCrop = ((Picture)picture.Element).BlipFill?.SourceRectangle != null;
            if (PowerPointImageSize.TryRead(picture.GetImageBytes(), out int width, out int height)) {
                if (!cover) {
                    double scale = Math.Min(box.Width / width, box.Height / height);
                    box = new PowerPointLayoutRect(box.X, box.Y, width * scale, height * scale);
                }
                ArrangeNativeFigure(writer.Slide, picture, box);
                if (cover && !hasCrop) picture.CropToFill(width, height, image.FocalX, image.FocalY);
            } else ArrangeNativeFigure(writer.Slide, picture, box);
            RenderAnnotations(context, writer, picture, image);
            return picture;
        }
    }
}
