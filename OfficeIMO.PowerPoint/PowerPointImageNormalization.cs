using System.IO;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointImageNormalization {
        internal static MemoryStream? NormalizeHeicForEmbedding(Stream image, OfficeImageFormat declaredFormat) {
            if (declaredFormat != OfficeImageFormat.Heic) return null;

            if (image.CanSeek) image.Position = 0;
            byte[] encoded = OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(
                image, new OfficeRasterDecodeOptions().MaximumEncodedBytes);
            byte[] normalized = OfficeImageReader.NormalizeForEmbedding(encoded, null, out OfficeImageInfo normalizedInfo);
            if (normalizedInfo.Format != OfficeImageFormat.Png)
                throw new InvalidDataException("The HEIC image could not be normalized to PNG for PowerPoint.");
            return new MemoryStream(normalized, writable: false);
        }
    }
}
