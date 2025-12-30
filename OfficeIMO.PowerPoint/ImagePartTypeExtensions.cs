using System;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlImagePartType = DocumentFormat.OpenXml.Packaging.ImagePartType;

namespace OfficeIMO.PowerPoint {
    internal static class ImagePartTypeExtensions {
        public static PartTypeInfo ToPartTypeInfo(this ImagePartType type) => type switch {
            ImagePartType.Png => OpenXmlImagePartType.Png,
            ImagePartType.Jpeg => OpenXmlImagePartType.Jpeg,
            ImagePartType.Gif => OpenXmlImagePartType.Gif,
            ImagePartType.Bmp => OpenXmlImagePartType.Bmp,
            ImagePartType.Tiff => OpenXmlImagePartType.Tiff,
            ImagePartType.Emf => OpenXmlImagePartType.Emf,
            ImagePartType.Wmf => OpenXmlImagePartType.Wmf,
            ImagePartType.Icon => OpenXmlImagePartType.Icon,
            ImagePartType.Pcx => OpenXmlImagePartType.Pcx,
            _ => throw new NotSupportedException($"Image type {type} is not supported."),
        };
    }
}
