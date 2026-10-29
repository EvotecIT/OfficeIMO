using System;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointPicture {
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
