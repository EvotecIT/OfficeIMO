using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static ExcelPartRef GetImagePart(ExcelPartRef drawingPart, byte[] bytes, string contentType, bool deduplicate) {
            var type = ToImagePartContentType(contentType);
            if (deduplicate) {
                foreach (ExcelPartRef candidate in drawingPart.PartsOfKind(ExcelPartKind.Image)) {
                    if (!string.Equals(candidate.ContentType, type.ContentType, StringComparison.OrdinalIgnoreCase)) continue;
                    using Stream source = candidate.OpenRead();
                    if (source.CanSeek && source.Length != bytes.Length) continue;
                    bool equal = true;
                    for (int i = 0; i < bytes.Length; i++) {
                        if (source.ReadByte() == bytes[i]) continue;
                        equal = false;
                        break;
                    }
                    if (equal && source.ReadByte() < 0) return candidate;
                }
            }
            ExcelPartRef part = drawingPart.AddImagePartWithExtension(type.ContentType, type.Extension);
            part.Feed(bytes);
            return part;
        }

        // An SVG picture stores its PNG fallback in the ordinary blip and the SVG in the Office 2016 blip extension,
        // which is what Excel writes and what older readers expect. The blip is made for <paramref name="blipFill"/>.
        private static OfficeXmlNode CreatePictureBlip(ExcelPartRef drawingPart, OfficeXmlNode blipFill, byte[] bytes, string contentType, bool deduplicate) {
            bool svg = string.Equals(ToImagePartContentType(contentType).ContentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase);
            OfficeXmlNode blip = blipFill.New(A, "blip");
            if (!svg || !OfficeIMO.Core.Internal.OfficeSvgPictureFallback.TryCreate(bytes, out byte[] png, out _, out _)) {
                blip.SetAttribute(R, "embed", drawingPart.GetIdOfPart(GetImagePart(drawingPart, bytes, contentType, deduplicate)));
                return blip;
            }

            string? fallbackId = drawingPart.GetIdOfPart(GetImagePart(drawingPart, png, "image/png", deduplicate));
            string? svgId = drawingPart.GetIdOfPart(GetImagePart(drawingPart, bytes, contentType, deduplicate));
            blip.SetAttribute(R, "embed", fallbackId);
            OfficeXmlNode extensions = blip.Append(blip.New(A, "extLst"));
            OfficeXmlNode extension = extensions.New(A, "ext");
            extension.SetAttribute("uri", OfficeIMO.Core.Internal.OfficeSvgPictureFallback.BlipExtensionUri);
            extension = extensions.Append(extension);
            extension.Append(NewSvgBlip(extension, svgId));
            return blip;
        }
    }
}
