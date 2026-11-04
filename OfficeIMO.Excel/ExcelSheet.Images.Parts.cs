using DocumentFormat.OpenXml.Packaging;

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
    }
}
