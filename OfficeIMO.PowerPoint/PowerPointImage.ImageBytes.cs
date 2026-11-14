using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointImage {
        /// <summary>
        ///     Returns a snapshot of the embedded image bytes for export adapters.
        /// </summary>
        public byte[] GetImageBytes() {
            return GetImageBytes(maximumBytes: null);
        }

        /// <summary>Returns the raster fallback of an SVG picture, for renderers that cannot draw its SVG.</summary>
        internal bool TryGetRasterFallback(int? maximumBytes, out byte[] bytes) {
            bytes = Array.Empty<byte>();
            if (GetSvgRelationshipId() == null || GetFallbackImagePart() is not ImagePart fallback) return false;
            using Stream source = fallback.GetStream(FileMode.Open, FileAccess.Read);
            bytes = maximumBytes.HasValue
                ? OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(source, maximumBytes.Value)
                : OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(source);
            return bytes.Length > 0;
        }

        /// <summary>Returns a bounded snapshot of the embedded image bytes.</summary>
        internal byte[] GetImageBytes(int? maximumBytes) {
            ImagePart imagePart = GetImagePart() ?? throw new InvalidOperationException("Picture has no embedded image part.");
            using Stream source = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            return maximumBytes.HasValue
                ? OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(
                    source, maximumBytes.Value)
                : OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(
                    source);
        }
    }
}
