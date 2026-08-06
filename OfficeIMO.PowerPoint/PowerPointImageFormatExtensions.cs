using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OpenXmlImagePartType = DocumentFormat.OpenXml.Packaging.ImagePartType;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointImageFormatExtensions {
        internal static PartTypeInfo ToPartTypeInfo(this OfficeImageFormat format) => format switch {
            OfficeImageFormat.Png => OpenXmlImagePartType.Png,
            OfficeImageFormat.Jpeg => OpenXmlImagePartType.Jpeg,
            OfficeImageFormat.Gif => OpenXmlImagePartType.Gif,
            OfficeImageFormat.Bmp => OpenXmlImagePartType.Bmp,
            OfficeImageFormat.Tiff => OpenXmlImagePartType.Tiff,
            OfficeImageFormat.Svg => OpenXmlImagePartType.Svg,
            OfficeImageFormat.Emf => OpenXmlImagePartType.Emf,
            OfficeImageFormat.Wmf => OpenXmlImagePartType.Wmf,
            OfficeImageFormat.Icon => OpenXmlImagePartType.Icon,
            OfficeImageFormat.Pcx => OpenXmlImagePartType.Pcx,
            _ => throw new NotSupportedException($"Image format {format} is not supported by PowerPoint image parts.")
        };

        internal static OfficeImageFormat FromImagePath(string imagePath) {
            OfficeImageFormat format = OfficeImageReader.FromExtension(imagePath);
            return format == OfficeImageFormat.Unknown ? OfficeImageFormat.Png : format.EnsurePowerPointImagePartSupport();
        }

        internal static OfficeImageFormat EnsurePowerPointImagePartSupport(this OfficeImageFormat format) {
            _ = format.ToPartTypeInfo();
            return format;
        }
    }
}
