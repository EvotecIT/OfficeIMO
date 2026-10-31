using System;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointPicture {
        /// <summary>Fills the current shape bounds using the image's stored pixel size.</summary>
        public void FitToBox() => FitToBox(crop: true);

        /// <summary>
        ///     Fits the image into the current shape bounds using the pixel size stored in the image itself.
        ///     With <paramref name="crop" /> the picture fills the box and the excess is cropped on the axis where
        ///     the image is longer; otherwise the frame shrinks to the image's aspect ratio, centred in the box.
        /// </summary>
        public void FitToBox(bool crop) {
            if (!TryGetPixelSize(out int width, out int height)) {
                throw new InvalidOperationException("The picture's pixel size could not be read from its image data.");
            }

            FitToBox(width, height, crop);
        }

        internal bool TryGetPixelSize(out int width, out int height) {
            width = 0;
            height = 0;
            try {
                return PowerPointImageSize.TryRead(GetImageBytes(), out width, out height);
            } catch (InvalidOperationException) {
                return false;
            }
        }

        /// <summary>Crops the source so it fills the frame, keeping the focal point (fractions 0..1) in view.</summary>
        internal void CropToFill(int pixelWidth, int pixelHeight, double focalX, double focalY) {
            double boxAspect = Width.ToEmus() / (double)Math.Max(1L, Height.ToEmus());
            double imageAspect = pixelWidth / (double)Math.Max(1, pixelHeight);
            double left = 0, top = 0, right = 0, bottom = 0;
            if (imageAspect > boxAspect) {
                double keep = boxAspect / imageAspect;
                double start = Clamp(focalX - keep / 2D, 0D, 1D - keep);
                left = start;
                right = 1D - keep - start;
            } else if (imageAspect < boxAspect) {
                double keep = imageAspect / boxAspect;
                double start = Clamp(focalY - keep / 2D, 0D, 1D - keep);
                top = start;
                bottom = 1D - keep - start;
            }

            Crop(left * 100D, top * 100D, right * 100D, bottom * 100D);
        }

        /// <summary>Shrinks the frame to the image's aspect ratio, centred in the current frame.</summary>
        internal void ContainInFrame() {
            if (!TryGetPixelSize(out int width, out int height)) return;
            double boxWidth = Width.ToPoints(), boxHeight = Height.ToPoints();
            if (boxWidth <= 0 || boxHeight <= 0) return;
            double imageAspect = width / (double)height, boxAspect = boxWidth / boxHeight;
            if (Math.Abs(imageAspect / boxAspect - 1D) < 0.005D) return;
            FitToBox(width, height, crop: false);
        }

        private static double Clamp(double value, double min, double max) => value < min ? min : value > max ? max : value;
    
        /// <summary>Shows the picture in greyscale at <paramref name="opacity" /> (logo walls).</summary>
        internal void SetMonochrome(double opacity) {
            DocumentFormat.OpenXml.Drawing.Blip? blip = Picture.BlipFill?.Blip;
            if (blip == null) return;
            blip.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.Grayscale>();
            blip.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.AlphaModulationFixed>();
            blip.PrependChild(new DocumentFormat.OpenXml.Drawing.Grayscale());
            blip.PrependChild(new DocumentFormat.OpenXml.Drawing.AlphaModulationFixed { Amount = (int)Math.Round(opacity * 100000D) });
        }
    }
}
