using System;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OpenXmlImagePartType = DocumentFormat.OpenXml.Packaging.ImagePartType;

namespace OfficeIMO.PowerPoint {
    internal static class ImagePartTypeExtensions {
        public static PartTypeInfo ToPartTypeInfo(this ImagePartType type) => type switch {
            ImagePartType.Png => OpenXmlImagePartType.Png,
            ImagePartType.Jpeg => OpenXmlImagePartType.Jpeg,
            ImagePartType.Gif => OpenXmlImagePartType.Gif,
            ImagePartType.Bmp => OpenXmlImagePartType.Bmp,
            ImagePartType.Tiff => OpenXmlImagePartType.Tiff,
            ImagePartType.Svg => OpenXmlImagePartType.Svg,
            ImagePartType.Emf => OpenXmlImagePartType.Emf,
            ImagePartType.Wmf => OpenXmlImagePartType.Wmf,
            ImagePartType.Icon => OpenXmlImagePartType.Icon,
            ImagePartType.Pcx => OpenXmlImagePartType.Pcx,
            _ => throw new NotSupportedException($"Image type {type} is not supported."),
        };

        public static ImagePartType FromImagePath(string imagePath) =>
            FromOfficeImageFormat(OfficeImageReader.FromExtension(imagePath));

        public static ImagePartType FromOfficeImageFormat(OfficeImageFormat format) => format switch {
            OfficeImageFormat.Jpeg => ImagePartType.Jpeg,
            OfficeImageFormat.Gif => ImagePartType.Gif,
            OfficeImageFormat.Bmp => ImagePartType.Bmp,
            OfficeImageFormat.Tiff => ImagePartType.Tiff,
            OfficeImageFormat.Svg => ImagePartType.Svg,
            OfficeImageFormat.Emf => ImagePartType.Emf,
            OfficeImageFormat.Wmf => ImagePartType.Wmf,
            OfficeImageFormat.Icon => ImagePartType.Icon,
            OfficeImageFormat.Pcx => ImagePartType.Pcx,
            OfficeImageFormat.Unknown => ImagePartType.Png,
            _ => throw new NotSupportedException($"Image format {format} is not supported by PowerPoint image parts.")
        };
    }
}
