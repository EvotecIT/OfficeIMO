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

        /// <summary>
        /// Parses Visio color cell values like "#RRGGBB" or "RGB(r,g,b)" to a Color.
        /// Returns black for unrecognized formats.
        /// </summary>
        public static Color FromVisioColor(string value) {
            if (string.IsNullOrWhiteSpace(value)) return Color.Black;
            value = value.Trim();
            if (value.StartsWith("#") && value.Length == 7) {
                var r = System.Convert.ToByte(value.Substring(1, 2), 16);
                var g = System.Convert.ToByte(value.Substring(3, 2), 16);
                var b = System.Convert.ToByte(value.Substring(5, 2), 16);
                return Color.FromRgb(r, g, b);
            }
            if (value.StartsWith("RGB(", StringComparison.OrdinalIgnoreCase) && value.EndsWith(")")) {
                var inner = value.Substring(4, value.Length - 5);
                var parts = inner.Split(',');
                if (parts.Length == 3 &&
                    byte.TryParse(parts[0], out var r) &&
                    byte.TryParse(parts[1], out var g) &&
                    byte.TryParse(parts[2], out var b)) {
                    return Color.FromRgb(r, g, b);
                }
            }
            return Color.Black;
        }
    }
}
