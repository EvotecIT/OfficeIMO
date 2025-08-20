using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Helper methods for Visio document generation.
    /// </summary>
    internal static class VisioHelpers {
        /// <summary>
        /// Converts a SixLabors.ImageSharp.Color to Visio RGB string format.
        /// </summary>
        /// <param name="color">The color to convert.</param>
        /// <returns>A string in the format "RGB(r,g,b)" where r, g, b are 0-255.</returns>
        public static string ToVisioRgb(this Color color) {
            var rgba = color.ToPixel<Rgba32>();
            return $"RGB({rgba.R},{rgba.G},{rgba.B})";
        }
        
        /// <summary>
        /// Converts a SixLabors.ImageSharp.Color to hex format for Visio.
        /// </summary>
        /// <param name="color">The color to convert.</param>
        /// <returns>A string in the format "#RRGGBB".</returns>
        public static string ToVisioHex(this Color color) {
            var rgba = color.ToPixel<Rgba32>();
            return $"#{rgba.R:X2}{rgba.G:X2}{rgba.B:X2}";
        }
    }
}