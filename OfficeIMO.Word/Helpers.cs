using System;
using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using SixLabors.ImageSharp.Formats;

namespace OfficeIMO.Word {
    public static partial class Helpers {
        /// <summary>
        /// Converts Color to Hex Color
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public static string ToHexColor(this SixLabors.ImageSharp.Color c) {
            return c.ToHex().Remove(6);
        }

        /// <summary>
        /// Opens up any file using assigned Application
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="open"></param>
        public static void Open(string filePath, bool open) {
            if (open) {
                ProcessStartInfo startInfo = new ProcessStartInfo(filePath) {
                    UseShellExecute = true
                };
                Process.Start(startInfo);
            }
        }

        /// <summary>
        /// Checks if file is locked/used by another process
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static bool IsFileLocked(this FileInfo file) {
            try {
                using (FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None)) {
                    stream.Close();
                }
            } catch (IOException) {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }

            //file is not locked
            return false;
        }

        /// <summary>
        /// Checks if file is locked/used by another process
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool IsFileLocked(this string fileName) {
            if (string.IsNullOrEmpty(fileName)) {
                return false;
            }
            if (!File.Exists(fileName)) {
                return false;
            }
            return IsFileLocked(new FileInfo(fileName));
        }

        internal static ImageCharacteristics GetImageCharacteristics(Stream imageStream) {
            using var img = SixLabors.ImageSharp.Image.Load(imageStream, out var imageFormat);
            imageStream.Position = 0;
            var type = ConvertToImagePartType(imageFormat);
            return new ImageCharacteristics(img.Width, img.Height, type);
        }

        private static CustomImagePartType ConvertToImagePartType(IImageFormat imageFormat) =>
            imageFormat.Name switch {
                "BMP" => CustomImagePartType.Bmp,
                "GIF" => CustomImagePartType.Gif,
                "JPEG" => CustomImagePartType.Jpeg,
                "PNG" => CustomImagePartType.Png,
                "TIFF" => CustomImagePartType.Tiff,
                _ => throw new ImageFormatNotSupportedException($"Image format not supported: {imageFormat.Name}.")
            };

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
    }

    internal record ImageCharacteristics(double Width, double Height, CustomImagePartType Type);
}
