using OfficeIMO.Drawing;
using System;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides various utility methods used throughout the library.
    /// </summary>
    internal static partial class Helpers {
        private const double PixelsPerInch = 96.0;

        internal static double ConvertPixelsToPoints(double pixels) {
            return pixels * 72 / PixelsPerInch;
        }

        internal static double ConvertPointsToPixels(double points) {
            return points * PixelsPerInch / 72;
        }
        /// <summary>
        /// Parses a color string that may or may not start with '#'.
        /// </summary>
        /// <param name="hex">Color value in hex without alpha or with '#'.</param>
        internal static OfficeIMO.Drawing.OfficeColor ParseColor(string hex) {
            if (string.IsNullOrEmpty(hex)) {
                throw new ArgumentException("Value cannot be null or empty.", nameof(hex));
            }
            if (!hex.StartsWith("#", StringComparison.Ordinal)) {
                hex = "#" + hex;
            }
            return OfficeIMO.Drawing.OfficeColor.Parse(hex);
        }

        /// <summary>
        /// Normalizes color input which may be a hex value or a named color.
        /// Hex values may be specified as three or six digits, with or without '#'.
        /// Returns an uppercase six-digit hex string without '#'.
        /// Throws <see cref="ArgumentException"/> if the value cannot be parsed
        /// as a valid color.
        /// </summary>
        internal static string? NormalizeColor(string? color) {
            if (string.IsNullOrEmpty(color)) {
                return null;
            }

            try {
                var parsed = OfficeIMO.Drawing.OfficeColor.Parse(color!);
                return parsed.ToRgbHex().ToUpperInvariant();
            } catch {
                if (!color!.StartsWith("#", StringComparison.Ordinal)) {
                    try {
                        var parsedHex = OfficeIMO.Drawing.OfficeColor.Parse("#" + color);
                        return parsedHex.ToRgbHex().ToUpperInvariant();
                    } catch {
                        // ignored so that ArgumentException below is thrown
                    }
                }
                throw new ArgumentException($"Invalid color value: {color}. Must be a valid hex color (3 or 6 characters) or named color.", nameof(color));
            }
        }

        /// <summary>
        /// Normalizes a required six-digit hexadecimal RGB value for DrawingML color attributes.
        /// </summary>
        internal static string NormalizeSixDigitRgb(string? color, string paramName) {
            string normalized = color?.Trim() ?? string.Empty;
            if (normalized.StartsWith("#", StringComparison.Ordinal)) normalized = normalized.Substring(1);
            if (normalized.Length != 6 || normalized.Any(character =>
                character is not (>= '0' and <= '9')
                && character is not (>= 'A' and <= 'F')
                && character is not (>= 'a' and <= 'f'))) {
                throw new ArgumentException("Color must be a six-digit hexadecimal RGB value.", paramName);
            }
            return normalized.ToUpperInvariant();
        }

        internal static string? NormalizeOpenXmlColor(string? color) {
            if (color == null) {
                return null;
            }

            string normalized = color.Replace("#", string.Empty);
            return string.Equals(normalized, "auto", StringComparison.OrdinalIgnoreCase)
                ? "auto"
                : normalized.ToUpperInvariant();
        }

        internal static TResult UseSeekableImageStream<TResult>(Stream imageStream, Func<Stream, TResult> action) {
            if (imageStream == null) {
                throw new ArgumentNullException(nameof(imageStream));
            }

            if (action == null) {
                throw new ArgumentNullException(nameof(action));
            }

            if (imageStream.CanSeek) {
                long originalPosition = imageStream.Position;
                imageStream.Position = 0;

                try {
                    return action(imageStream);
                } finally {
                    imageStream.Position = originalPosition;
                }
            }

            using var copy = new MemoryStream();
            imageStream.CopyTo(copy);
            copy.Position = 0;
            return action(copy);
        }

        internal static ImageCharacteristics GetImageCharacteristics(Stream imageStream, string? fileName = null) {
            return UseSeekableImageStream(imageStream, seekableImageStream => GetImageCharacteristicsCore(seekableImageStream, fileName));
        }

        private static ImageCharacteristics GetImageCharacteristicsCore(Stream imageStream, string? fileName) {
            if (OfficeImageReader.TryIdentify(imageStream, fileName, out var imageInfo)) {
                return new ImageCharacteristics(imageInfo.Width, imageInfo.Height, EnsureSupportedImageFormat(imageInfo.Format));
            }

            return new ImageCharacteristics(0, 0, OfficeImageFormat.Png);
        }

        private static OfficeImageFormat EnsureSupportedImageFormat(OfficeImageFormat imageFormat) =>
            imageFormat is OfficeImageFormat.Bmp or
                OfficeImageFormat.Gif or
                OfficeImageFormat.Jpeg or
                OfficeImageFormat.Png or
                OfficeImageFormat.Tiff or
                OfficeImageFormat.Emf or
                OfficeImageFormat.Wmf or
                OfficeImageFormat.Svg
                ? imageFormat
                : throw new NotSupportedException($"Word image parts do not support {imageFormat} images.");

        /// <summary>
        /// Converts centimeters to EMUs and returns int value
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal static int? ConvertCentimetersToEmus(double value) {
            int emus = (int)(value * 360000);
            return emus;
        }

        /// <summary>
        /// Converts centimeters to EMUs and returns int64 value
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal static Int64 ConvertCentimetersToEmusInt64(double value) {
            Int64 emus = (Int64)(value * 360000);
            return emus;
        }

        /// <summary>
        /// Converts EMUs to centimeters
        /// </summary>
        /// <param name="emusValue"></param>
        /// <returns></returns>
        internal static double? ConvertEmusToCentimeters(int emusValue) {
            double centimeters = (double)((double)emusValue / (double)360000);
            return centimeters;
        }

        /// <summary>
        /// Converts EMUs to centimeters
        /// </summary>
        /// <param name="emusValue"></param>
        /// <returns></returns>
        internal static double ConvertEmusToCentimeters(Int64 emusValue) {
            double centimeters = (double)((double)emusValue / (double)360000);
            return centimeters;
        }

        /// <summary>
        /// Converts twips to centimeters
        /// </summary>
        /// <param name="twipsValue"></param>
        /// <returns></returns>
        internal static double ConvertTwipsToCentimeters(int twipsValue) {
            double centimeters = twipsValue / 567.0;
            return Math.Round(centimeters, 2);
        }

        /// <summary>
        /// Converts twips to centimeters
        /// </summary>
        /// <param name="twipsValue"></param>
        /// <returns></returns>
        internal static double ConvertTwipsToCentimeters(UInt32 twipsValue) {
            double centimeters = twipsValue / 567.0;
            return Math.Round(centimeters, 2);
        }

        /// <summary>
        /// Converts centimeters to twips
        /// </summary>
        /// <param name="cmValue"></param>
        /// <returns></returns>
        internal static int ConvertCentimetersToTwips(double cmValue) {
            int twips = (int)Math.Round(cmValue * 567.0);
            return twips;
        }

        /// <summary>
        /// Converts centimeters to twips
        /// </summary>
        /// <param name="cmValue"></param>
        /// <returns></returns>
        internal static UInt32 ConvertCentimetersToTwipsUInt32(double cmValue) {
            UInt32 twips = (UInt32)Math.Round(cmValue * 567.0);
            return twips;
        }

        /// <summary>
        /// Converts centimeters to twentieths of a point
        /// </summary>
        /// <param name="cm"></param>
        /// <returns></returns>
        internal static double ConvertCentimetersToTwentiethsOfPoint(double cm) {
            double inches = cm / 2.54;
            double points = inches * 72;
            double twentiethsOfPoint = points * 20;
            return twentiethsOfPoint;
        }

        internal static double ConvertTwentiethsOfPointToCentimeters(double twentiethsOfPoint) {
            double points = twentiethsOfPoint / 20;
            double centimeters = (points / 72) * 2.54;
            return centimeters;
        }

        /// <summary>
        /// Converts centimeters to points
        /// </summary>
        /// <param name="cm"></param>
        /// <returns></returns>
        internal static double ConvertCentimetersToPoints(double cm) {
            double inches = cm / 2.54;
            double points = inches * 72;
            return points;
        }

        /// <summary>
        /// Converts the points to centimeters.
        /// </summary>
        /// <param name="points">The points.</param>
        /// <returns></returns>
        internal static double ConvertPointsToCentimeters(double points) {
            double centimeters = (points / 72) * 2.54;
            return centimeters;
        }

        /// <summary>
        /// Converts twips (twentieths of a point) to points
        /// </summary>
        /// <param name="twipsValue">The value in twips.</param>
        /// <returns>Points value.</returns>
        internal static double ConvertTwipsToPoints(int twipsValue) {
            double points = twipsValue / 20.0;
            return Math.Round(points, 2);
        }

        /// <summary>
        /// Converts twips (twentieths of a point) to points
        /// </summary>
        /// <param name="twipsValue">The value in twips.</param>
        /// <returns>Points value.</returns>
        internal static double ConvertTwipsToPoints(UInt32 twipsValue) {
            double points = twipsValue / 20.0;
            return Math.Round(points, 2);
        }

        /// <summary>
        /// Converts points to twips (twentieths of a point)
        /// </summary>
        /// <param name="points">The points value.</param>
        /// <returns>Twips value.</returns>
        internal static int ConvertPointsToTwips(double points) {
            int twips = (int)Math.Round(points * 20.0);
            return twips;
        }

        /// <summary>
        /// Converts points to twips (twentieths of a point)
        /// </summary>
        /// <param name="points">The points value.</param>
        /// <returns>Twips value.</returns>
        internal static UInt32 ConvertPointsToTwipsUInt32(double points) {
            UInt32 twips = (UInt32)Math.Round(points * 20.0);
            return twips;
        }

        /// <summary>
        /// Converts EMUs to points
        /// </summary>
        /// <param name="emusValue">EMUs value.</param>
        /// <returns>Points value.</returns>
        internal static double ConvertEmusToPoints(Int64 emusValue) {
            double points = emusValue / 12700.0;
            return points;
        }

        /// <summary>
        /// Converts points to EMUs
        /// </summary>
        /// <param name="points">Points value.</param>
        /// <returns>EMUs value.</returns>
        internal static Int64 ConvertPointsToEmusInt64(double points) {
            Int64 emus = (Int64)(points * 12700.0);
            return emus;
        }
    }

    internal record ImageCharacteristics(double Width, double Height, OfficeImageFormat Type);
}