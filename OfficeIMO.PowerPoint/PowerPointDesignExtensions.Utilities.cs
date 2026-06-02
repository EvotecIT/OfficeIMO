using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public static partial class PowerPointDesignExtensions {
        internal static PowerPointTextBox AddText(PowerPointSlide slide, string text, double leftCm, double topCm,
            double widthCm, double heightCm, int fontSize, string color, string fontName, bool bold = false) {
            PowerPointTextBox box = slide.AddTextBoxCm(text, leftCm, topCm, widthCm, heightCm);
            box.SetTextMarginsCm(0, 0, 0, 0);
            box.FontName = fontName;
            box.FontSize = fontSize;
            box.Color = color;
            box.Bold = bold;
            box.TextAutoFit = PowerPointTextAutoFit.Normal;
            return box;
        }

        private static PowerPointPicture? AddPictureIfExists(PowerPointSlide slide, string imagePath,
            double leftCm, double topCm, double widthCm, double heightCm, bool crop) {
            if (!File.Exists(imagePath)) {
                return null;
            }

            PowerPointPicture picture = slide.AddPictureCm(imagePath, leftCm, topCm, widthCm, heightCm);
            if (crop && TryGetImageDimensions(imagePath, out double imageWidth, out double imageHeight)) {
                picture.FitToBox(imageWidth, imageHeight, crop: true);
            }
            return picture;
        }

        private static bool TryGetImageDimensions(string imagePath, out double width, out double height) {
            width = 0;
            height = 0;

            using FileStream stream = File.OpenRead(imagePath);
            if (TryGetPngDimensions(stream, out width, out height)) {
                return true;
            }

            stream.Position = 0;
            return TryGetJpegDimensions(stream, out width, out height);
        }

        private static bool TryGetPngDimensions(Stream stream, out double width, out double height) {
            width = 0;
            height = 0;

            byte[] header = new byte[24];
            if (stream.Read(header, 0, header.Length) != header.Length) {
                return false;
            }

            byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
            for (int i = 0; i < signature.Length; i++) {
                if (header[i] != signature[i]) {
                    return false;
                }
            }

            width = ReadBigEndianInt32(header, 16);
            height = ReadBigEndianInt32(header, 20);
            return width > 0 && height > 0;
        }

        private static bool TryGetJpegDimensions(Stream stream, out double width, out double height) {
            width = 0;
            height = 0;

            int first = stream.ReadByte();
            int second = stream.ReadByte();
            if (first != 0xFF || second != 0xD8) {
                return false;
            }

            while (stream.Position < stream.Length) {
                int markerPrefix;
                do {
                    markerPrefix = stream.ReadByte();
                    if (markerPrefix < 0) {
                        return false;
                    }
                } while (markerPrefix != 0xFF);

                int marker;
                do {
                    marker = stream.ReadByte();
                    if (marker < 0) {
                        return false;
                    }
                } while (marker == 0xFF);

                if (marker is 0xD8 or 0xD9) {
                    continue;
                }

                int segmentLength = ReadBigEndianUInt16(stream);
                if (segmentLength < 2 || stream.Position + segmentLength - 2 > stream.Length) {
                    return false;
                }

                if (IsJpegStartOfFrame(marker)) {
                    stream.ReadByte();
                    height = ReadBigEndianUInt16(stream);
                    width = ReadBigEndianUInt16(stream);
                    return width > 0 && height > 0;
                }

                stream.Position += segmentLength - 2;
            }

            return false;
        }

        private static bool IsJpegStartOfFrame(int marker) {
            return marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or
                0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;
        }

        private static int ReadBigEndianInt32(byte[] bytes, int offset) {
            return (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];
        }

        private static int ReadBigEndianUInt16(Stream stream) {
            int high = stream.ReadByte();
            int low = stream.ReadByte();
            if (high < 0 || low < 0) {
                return -1;
            }

            return (high << 8) | low;
        }

        private static void CenterText(PowerPointTextBox textBox) {
            foreach (PowerPointParagraph paragraph in textBox.Paragraphs) {
                paragraph.Alignment = A.TextAlignmentTypeValues.Center;
            }
            textBox.TextVerticalAlignment = A.TextAnchoringTypeValues.Center;
        }

        private static void RightAlignText(PowerPointTextBox textBox) {
            foreach (PowerPointParagraph paragraph in textBox.Paragraphs) {
                paragraph.Alignment = A.TextAlignmentTypeValues.Right;
            }
        }

        private static string GetAccent(PowerPointDesignTheme theme, int index) {
            string[] colors = {
                theme.AccentColor,
                theme.Accent3Color,
                theme.Accent2Color,
                theme.AccentDarkColor,
                theme.WarningColor
            };
            return colors[index % colors.Length];
        }
    }
}
