using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Checks whether image bytes can be embedded by the first-party PDF writer.
    /// </summary>
    public static bool TryValidateImageBytes(byte[] data, out OfficeImageInfo? imageInfo, out string? unsupportedReason) {
        imageInfo = null;
        unsupportedReason = null;
        try {
            imageInfo = ValidateImageBytes(data);
            return true;
        } catch (NotSupportedException ex) {
            unsupportedReason = ex.Message;
            return false;
        } catch (ArgumentException ex) {
            unsupportedReason = ex.Message;
            return false;
        }
    }

    internal static OfficeImageInfo ValidateImageBytes(byte[] data) {
        if (OfficeImageReader.TryIdentify(data, null, out var info)) {
            if (info.Format == OfficeImageFormat.Jpeg) {
                return info;
            }

            if (info.Format == OfficeImageFormat.Png) {
                string? unsupportedReason;
                if (PdfWriter.TryGetPngImageData(data, out _, out unsupportedReason)) {
                    return info;
                }

                throw new NotSupportedException(
                    "PdfDocument.Image currently supports JPEG and grayscale/grayscale-alpha/indexed-color/RGB/RGBA PNG image bytes only, including Adam7-interlaced PNGs and supported 16-bit grayscale/grayscale-alpha/RGB/RGBA PNG payloads. " +
                    unsupportedReason);
            } else {
                throw new NotSupportedException(
                    $"PdfDocument.Image currently supports JPEG and grayscale/grayscale-alpha/indexed-color/RGB/RGBA PNG image bytes only, including Adam7-interlaced PNGs and supported 16-bit grayscale/grayscale-alpha/RGB/RGBA PNG payloads. Detected {info.Format} ({info.MimeType}).");
            }
        }

        if (!LooksLikeJpeg(data)) {
            System.Diagnostics.Trace.TraceWarning("PdfDocument.Image: Provided bytes do not appear to be JPEG encoded.");
        }

        return new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
    }

    private static bool LooksLikeJpeg(byte[] data) {
        if (data.Length < 4)
            return false;

        return data[0] == 0xFF && data[1] == 0xD8 && data[data.Length - 2] == 0xFF && data[data.Length - 1] == 0xD9;
    }
}
