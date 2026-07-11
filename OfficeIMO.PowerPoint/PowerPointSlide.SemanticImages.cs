using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds a semantic image with deterministic fit/fill behavior, focal-point crop, and alternative text.
        /// </summary>
        public PowerPointPicture AddPicture(PowerPointImageAsset asset, PowerPointLayoutBox bounds) {
            if (asset == null) throw new ArgumentNullException(nameof(asset));
            asset.Validate();
            PowerPointPicture picture = AddPicture(asset.Path, bounds);
            picture.AltText = asset.AlternativeText;
            picture.Name = string.IsNullOrWhiteSpace(asset.Caption) ? "Semantic Image" : asset.Caption!;
            if (asset.Placement == PowerPointImagePlacement.Stretch) return picture;

            OfficeImageInfo info = OfficeImageReader.Identify(asset.Path);
            if (info.Width <= 0 || info.Height <= 0) return picture;
            if (asset.Placement == PowerPointImagePlacement.Fit) {
                picture.FitToBox(info.Width, info.Height, crop: false);
                return picture;
            }

            ApplyFocalCrop(picture, info.Width, info.Height, bounds.Width, bounds.Height,
                asset.FocalX, asset.FocalY);
            return picture;
        }

        private static void ApplyFocalCrop(PowerPointPicture picture, double imageWidth, double imageHeight,
            double boxWidth, double boxHeight, double focalX, double focalY) {
            double imageAspect = imageWidth / imageHeight;
            double boxAspect = boxWidth / boxHeight;
            double left = 0D;
            double right = 0D;
            double top = 0D;
            double bottom = 0D;
            if (imageAspect > boxAspect) {
                double visible = boxAspect / imageAspect;
                left = Clamp(focalX - visible / 2D, 0D, 1D - visible);
                right = 1D - visible - left;
            } else if (imageAspect < boxAspect) {
                double visible = imageAspect / boxAspect;
                top = Clamp(focalY - visible / 2D, 0D, 1D - visible);
                bottom = 1D - visible - top;
            }
            picture.Crop(left * 100D, top * 100D, right * 100D, bottom * 100D);
        }

        private static double Clamp(double value, double minimum, double maximum) =>
            Math.Max(minimum, Math.Min(maximum, value));
    }
}
