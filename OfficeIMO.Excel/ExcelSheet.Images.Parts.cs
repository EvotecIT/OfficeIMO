using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static ImagePart GetImagePart(DrawingsPart drawingPart, byte[] bytes, string contentType, bool deduplicate) {
            var type = ToImagePartType(contentType);
            if (deduplicate) {
                foreach (ImagePart candidate in drawingPart.ImageParts) {
                    if (!string.Equals(candidate.ContentType, type.ContentType, StringComparison.OrdinalIgnoreCase)) continue;
                    using Stream source = candidate.GetStream(FileMode.Open, FileAccess.Read);
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
            ImagePart part = drawingPart.AddImagePart(type);
            using (var source = new MemoryStream(bytes, writable: false)) part.FeedData(source);
            return part;
        }

        // An SVG picture stores its PNG fallback in the ordinary blip and the SVG in the Office 2016 blip extension,
        // which is what Excel writes and what older readers expect.
        private static A.Blip CreatePictureBlip(DrawingsPart drawingPart, byte[] bytes, string contentType, bool deduplicate) {
            bool svg = string.Equals(ToImagePartType(contentType).ContentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase);
            if (!svg || !OfficeIMO.Core.Internal.OfficeSvgPictureFallback.TryCreate(bytes, out byte[] png, out _, out _)) {
                return new A.Blip { Embed = drawingPart.GetIdOfPart(GetImagePart(drawingPart, bytes, contentType, deduplicate)) };
            }

            string fallbackId = drawingPart.GetIdOfPart(GetImagePart(drawingPart, png, "image/png", deduplicate));
            string svgId = drawingPart.GetIdOfPart(GetImagePart(drawingPart, bytes, contentType, deduplicate));
            return new A.Blip(new A.BlipExtensionList(new A.BlipExtension(
                new DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip { Embed = svgId }) {
                Uri = OfficeIMO.Core.Internal.OfficeSvgPictureFallback.BlipExtensionUri
            })) { Embed = fallbackId };
        }
    }
}
