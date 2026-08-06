using System;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OpenXmlImagePartType = DocumentFormat.OpenXml.Packaging.ImagePartType;

namespace OfficeIMO.PowerPoint {
    internal static class ImagePartTypeExtensions {
        public static PartTypeInfo ToPartTypeInfo(this PowerPointImagePartType type) => type switch {
            PowerPointImagePartType.Png => OpenXmlImagePartType.Png,
            PowerPointImagePartType.Jpeg => OpenXmlImagePartType.Jpeg,
            PowerPointImagePartType.Gif => OpenXmlImagePartType.Gif,
            PowerPointImagePartType.Bmp => OpenXmlImagePartType.Bmp,
            PowerPointImagePartType.Tiff => OpenXmlImagePartType.Tiff,
            PowerPointImagePartType.Svg => OpenXmlImagePartType.Svg,
            PowerPointImagePartType.Emf => OpenXmlImagePartType.Emf,
            PowerPointImagePartType.Wmf => OpenXmlImagePartType.Wmf,
            PowerPointImagePartType.Icon => OpenXmlImagePartType.Icon,
            PowerPointImagePartType.Pcx => OpenXmlImagePartType.Pcx,
            _ => throw new NotSupportedException($"Image type {type} is not supported."),
        };

        public static PowerPointImagePartType FromImagePath(string imagePath) =>
            FromOfficeImageFormat(OfficeImageReader.FromExtension(imagePath));

        public static PowerPointImagePartType FromOfficeImageFormat(OfficeImageFormat format) => format switch {
            OfficeImageFormat.Png => PowerPointImagePartType.Png,
            OfficeImageFormat.Jpeg => PowerPointImagePartType.Jpeg,
            OfficeImageFormat.Gif => PowerPointImagePartType.Gif,
            OfficeImageFormat.Bmp => PowerPointImagePartType.Bmp,
            OfficeImageFormat.Tiff => PowerPointImagePartType.Tiff,
            OfficeImageFormat.Svg => PowerPointImagePartType.Svg,
            OfficeImageFormat.Emf => PowerPointImagePartType.Emf,
            OfficeImageFormat.Wmf => PowerPointImagePartType.Wmf,
            OfficeImageFormat.Icon => PowerPointImagePartType.Icon,
            OfficeImageFormat.Pcx => PowerPointImagePartType.Pcx,
            OfficeImageFormat.Unknown => PowerPointImagePartType.Png,
            _ => throw new NotSupportedException($"Image format {format} is not supported by PowerPoint image parts.")
        };
    }
}
