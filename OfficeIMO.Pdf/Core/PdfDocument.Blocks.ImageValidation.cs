using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    private const string SupportedImageMessage =
        "PdfDocument.Image currently supports JPEG and grayscale/grayscale-alpha/indexed-color/RGB/RGBA PNG image bytes only, including Adam7-interlaced PNGs and supported 16-bit grayscale/grayscale-alpha/RGB/RGBA PNG payloads.";

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

                throw new NotSupportedException(SupportedImageMessage + " " + unsupportedReason);
            } else {
                throw new NotSupportedException(
                    $"{SupportedImageMessage} Detected {info.Format} ({info.MimeType}).");
            }
        }

        if (LooksLikePng(data)) {
            string? unsupportedReason;
            if (PdfWriter.TryGetPngImageData(data, out var image, out unsupportedReason)) {
                return new OfficeImageInfo(OfficeImageFormat.Png, image.PixelWidth, image.PixelHeight);
            }

            throw new NotSupportedException(SupportedImageMessage + " " + unsupportedReason);
        }

        if (!LooksLikeJpeg(data)) {
            System.Diagnostics.Trace.TraceWarning("PdfDocument.Image: Provided bytes do not appear to be JPEG encoded.");
        }

        return new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
    }

    private static bool LooksLikePng(byte[] data) =>
        data.Length >= 8 &&
        data[0] == 137 &&
        data[1] == 80 &&
        data[2] == 78 &&
        data[3] == 71 &&
        data[4] == 13 &&
        data[5] == 10 &&
        data[6] == 26 &&
        data[7] == 10;

    private static bool LooksLikeJpeg(byte[] data) {
        if (data.Length < 4)
            return false;

        return data[0] == 0xFF && data[1] == 0xD8 && data[data.Length - 2] == 0xFF && data[data.Length - 1] == 0xD9;
    }
}
