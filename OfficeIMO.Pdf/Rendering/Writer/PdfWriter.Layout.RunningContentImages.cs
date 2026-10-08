using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>
    /// Owns invariant image payloads for one generation, including pagination stabilization.
    /// Page placements stay independent; encoded data and prepared streams remain immutable.
    /// </summary>
    private sealed class RunningContentImageAssets {
        // Bound all retained running-story payloads, including decoded streams and masks,
        // even when no generated-page/output limit was supplied by the caller.
        private const long MaximumBytes = PdfImageInput.DefaultMaximumEncodedBytes;
        private readonly Dictionary<string, List<PageImage>> assets = new(StringComparer.Ordinal);
        private long retainedBytes;

        internal void Share(PageImage image, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            string key;
#if NET6_0_OR_GREATER
            key = Convert.ToBase64String(SHA256.HashData(image.Data));
#else
            using (var sha = SHA256.Create()) key = Convert.ToBase64String(sha.ComputeHash(image.Data));
#endif
            cancellationToken.ThrowIfCancellationRequested();
            if (assets.TryGetValue(key, out List<PageImage>? matches)) {
                foreach (PageImage asset in matches) {
                    if (image.Info.Format == asset.Info.Format && image.Info.Width == asset.Info.Width &&
                        image.Info.Height == asset.Info.Height && BytesEqual(image.Data, asset.Data) &&
                        (image.PreparedStream == null ? asset.PreparedStream == null :
                            asset.PreparedStream != null && SameImageStream(image.PreparedStream, asset.PreparedStream))) {
                        image.Data = asset.Data;
                        image.PreparedStream = asset.PreparedStream;
                        return;
                    }
                }
            }

            long bytes = GetRetainedBytes(image);
            if (bytes > MaximumBytes - retainedBytes)
                throw new InvalidDataException("Running PDF content exceeded the 128 MiB retained image asset limit.");
            retainedBytes += bytes;
            if (matches == null) assets.Add(key, matches = new List<PageImage>());
            // Retain only payloads and immutable metadata, not a page's mutable placement.
            matches.Add(new PageImage { Data = image.Data, Info = image.Info, PreparedStream = image.PreparedStream });
        }

        private static long GetRetainedBytes(PageImage image) {
            long bytes = image.Data.LongLength;
            if (image.PreparedStream != null) {
                if (!ReferenceEquals(image.Data, image.PreparedStream.Data)) bytes += image.PreparedStream.Data.LongLength;
                if (image.PreparedStream.SoftMask != null &&
                    !ReferenceEquals(image.Data, image.PreparedStream.SoftMask.Data) &&
                    !ReferenceEquals(image.PreparedStream.Data, image.PreparedStream.SoftMask.Data))
                    bytes += image.PreparedStream.SoftMask.Data.LongLength;
            }
            return bytes;
        }
    }
}
