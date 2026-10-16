using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Normalizes orientation once for every composition surface, before layout sees dimensions.</summary>
    private static PreparedImage PrepareOrientedImage(byte[] source, OfficeImageInfo info,
        OfficeImageMetadataSnapshot metadata, OfficeImageOrientation orientation, CancellationToken cancellationToken) {
        if (!OfficeImagePdfCompatibility.TryValidateTranscodeDimensions(info,
                OfficeImagePdfCompatibility.DefaultMaximumTranscodePixels, out string? limitReason)) {
            throw new NotSupportedException(SupportedImageMessage + " " + limitReason);
        }
        byte[] png;
        bool colorManaged = (metadata.Kinds & OfficeImageMetadataKinds.Icc) != 0;
        if (colorManaged) {
            // Orientation needs decoded pixels. Use the existing color-managed decoder here;
            // copying a gray/CMYK profile onto its RGB output would describe different samples.
            byte[]? profileBytes = metadata.Icc;
            if (profileBytes == null ||
                !PdfWriter.TryGetIccParseAllocationUpperBound(profileBytes.LongLength, out long parseBytes) ||
                source.LongLength + profileBytes.LongLength + parseBytes + (long)info.Width * info.Height * 8L > OfficeRasterGuards.MaximumDecodedBytes ||
                !OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile) || profile == null ||
                !OfficeIccRasterConverter.TryDecodeToSrgb(source, profile,
                    new OfficeRasterDecodeOptions { CancellationToken = cancellationToken }, out OfficeRasterImage? raster) || raster == null) {
                throw new NotSupportedException("The oriented image ICC profile cannot be normalized safely for PDF composition.");
            }
            bool swapAxes = orientation >= OfficeImageOrientation.Transpose;
            using var output = new MemoryStream();
            OfficePngWriter.EncodeRgbTo(raster, output, new OfficePngEncodeOptions {
                DpiX = swapAxes ? info.DpiY : info.DpiX,
                DpiY = swapAxes ? info.DpiX : info.DpiY
            }, cancellationToken);
            png = output.ToArray();
        } else if (!OfficeImageOrientationNormalizer.TryNormalizeToPng(source, true, cancellationToken,
                       out png, out _)) {
            throw new NotSupportedException("The image has embedded orientation metadata that could not be normalized safely for PDF composition.");
        }
        if (colorManaged) {
            // The pixels now describe sRGB. Preserve that calibration in both the
            // normalized source and its PDF resource, including later optimization.
            OfficeImageMetadata normalizedMetadata = OfficeImageMetadata.Read(png, cancellationToken);
            normalizedMetadata.IccProfile = PdfIccProfiles.SrgbIec6196621;
            png = OfficeImageMetadata.Apply(png, normalizedMetadata, cancellationToken);
        }
        PreparedImage prepared = PrepareImageBytesCore(png, cancellationToken);
        return new PreparedImage(prepared.Data, prepared.Info, info.Format, wasTranscoded: true, prepared.PreparedStream);
    }
}
