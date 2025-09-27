using System;
using System.Globalization;
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
            return TryParseVisioColor(value, out Color color) ? color : Color.Black;
        }

        /// <summary>
        /// Attempts to parse a Visio color string (handles hex, RGB, and guarded values).
        /// </summary>
        public static bool TryParseVisioColor(string value, out Color color) {
            color = Color.Black;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            string normalized = NormalizeVisioColor(value);

            if (normalized.StartsWith("#", StringComparison.Ordinal) && normalized.Length == 7) {
                try {
                    byte r = Convert.ToByte(normalized.Substring(1, 2), 16);
                    byte g = Convert.ToByte(normalized.Substring(3, 2), 16);
                    byte b = Convert.ToByte(normalized.Substring(5, 2), 16);
                    color = Color.FromRgb(r, g, b);
                    return true;
                } catch (FormatException) {
                    return false;
                }
            }

            if (normalized.StartsWith("0x", StringComparison.OrdinalIgnoreCase) && normalized.Length >= 3) {
                if (int.TryParse(normalized.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int hex)) {
                    byte r = (byte)((hex >> 16) & 0xFF);
                    byte g = (byte)((hex >> 8) & 0xFF);
                    byte b = (byte)(hex & 0xFF);
                    color = Color.FromRgb(r, g, b);
                    return true;
                }
            }

            int rgbIndex = normalized.IndexOf("RGB(", StringComparison.OrdinalIgnoreCase);
            if (rgbIndex >= 0) {
                int endIndex = normalized.IndexOf(')', rgbIndex);
                if (endIndex > rgbIndex) {
                    string inner = normalized.Substring(rgbIndex + 4, endIndex - rgbIndex - 4);
                    string[] parts = inner.Split(',');
                    if (parts.Length == 3) {
                        if (TryParseComponent(parts[0], out byte r) &&
                            TryParseComponent(parts[1], out byte g) &&
                            TryParseComponent(parts[2], out byte b)) {
                            color = Color.FromRgb(r, g, b);
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        private static string NormalizeVisioColor(string value) {
            string trimmed = value.Trim();
            bool changed;
            do {
                changed = false;
                trimmed = trimmed.Trim();
                string? unwrapped = TryUnwrap(trimmed, "THEMEGUARD");
                if (unwrapped != null) {
                    trimmed = unwrapped;
                    changed = true;
                    continue;
                }
                unwrapped = TryUnwrap(trimmed, "GUARD");
                if (unwrapped != null) {
                    trimmed = unwrapped;
                    changed = true;
                }
            } while (changed);

            // If both a result and formula are concatenated (e.g. "#000000;RGB(...)") pick the first hex or RGB fragment.
            int hashIndex = trimmed.IndexOf('#');
            if (hashIndex >= 0 && trimmed.Length >= hashIndex + 7) {
                return trimmed.Substring(hashIndex, 7);
            }

            int rgbIndex = trimmed.IndexOf("RGB(", StringComparison.OrdinalIgnoreCase);
            if (rgbIndex >= 0) {
                int endIndex = trimmed.IndexOf(')', rgbIndex);
                if (endIndex > rgbIndex) {
                    return trimmed.Substring(rgbIndex, endIndex - rgbIndex + 1);
                }
            }

            return trimmed;
        }

        private static string? TryUnwrap(string value, string wrapper) {
            if (value.StartsWith(wrapper + "(", StringComparison.OrdinalIgnoreCase) && value.EndsWith(")", StringComparison.Ordinal)) {
                return value.Substring(wrapper.Length + 1, value.Length - wrapper.Length - 2);
            }
            return null;
        }

        private static bool TryParseComponent(string component, out byte result) {
            result = 0;
            if (!double.TryParse(component.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) {
                return false;
            }

            if (value <= 1.0 && value >= 0.0) {
                value *= 255.0;
            }

            value = Math.Max(0.0, Math.Min(255.0, value));
            result = (byte)Math.Round(value, MidpointRounding.AwayFromZero);
            return true;
        }
    }
}
