using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static OfficeImageInfo? NormalizeHeicImageForEmbedding(ref byte[] bytes, ref string contentType,
            CancellationToken cancellationToken = default) {
            if (!OfficeImageReader.TryIdentifyByContent(bytes, null, cancellationToken, out OfficeImageInfo info) ||
                info.Format != OfficeImageFormat.Heic) return null;

            bytes = OfficeImageReader.NormalizeForEmbedding(bytes,
                new OfficeRasterDecodeOptions { CancellationToken = cancellationToken }, out OfficeImageInfo normalizedInfo);
            contentType = normalizedInfo.MimeType;
            return normalizedInfo;
        }
    }
}
