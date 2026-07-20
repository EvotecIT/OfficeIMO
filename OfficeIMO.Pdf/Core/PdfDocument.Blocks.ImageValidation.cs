using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    private const string SupportedImageMessage =
        "PdfDocument.Image accepts JPEG and the raster formats decoded by OfficeIMO.Drawing. JPEG and writer-safe PNG payloads are embedded directly; other supported raster payloads are normalized to PNG once before PDF serialization.";

    internal readonly struct PreparedImage {
        internal PreparedImage(byte[] data, OfficeImageInfo info, OfficeImageFormat sourceFormat, bool wasTranscoded) {
            Data = data;
            Info = info;
            SourceFormat = sourceFormat;
            WasTranscoded = wasTranscoded;
        }

        internal byte[] Data { get; }
        internal OfficeImageInfo Info { get; }
        internal OfficeImageFormat SourceFormat { get; }
        internal bool WasTranscoded { get; }
    }

    /// <summary>
    /// Checks whether image bytes can be embedded by the first-party PDF writer.
    /// </summary>
    public static bool TryValidateImageBytes(byte[] data, out OfficeImageInfo? imageInfo, out string? unsupportedReason) {
        bool prepared = TryPrepareImageBytes(data, out _, out imageInfo, out _, out unsupportedReason);
        return prepared;
    }

    /// <summary>
    /// Prepares source image bytes for first-party PDF embedding. Writer-safe JPEG and PNG data is retained;
    /// other raster formats supported by <see cref="OfficeRasterImageDecoder"/> are normalized to PNG.
    /// </summary>
    public static bool TryPrepareImageBytes(
        byte[] data,
        out byte[] preparedBytes,
        out OfficeImageInfo? imageInfo,
        out bool wasTranscoded,
        out string? unsupportedReason) {
        preparedBytes = System.Array.Empty<byte>();
        imageInfo = null;
        wasTranscoded = false;
        unsupportedReason = null;
        try {
            PreparedImage prepared = PrepareImageBytes(data);
            preparedBytes = (byte[])prepared.Data.Clone();
            imageInfo = prepared.Info;
            wasTranscoded = prepared.WasTranscoded;
            return true;
        } catch (NotSupportedException ex) {
            unsupportedReason = ex.Message;
            return false;
        } catch (ArgumentException ex) {
            unsupportedReason = ex.Message;
            return false;
        }
    }

    internal static OfficeImageInfo ValidateImageBytes(byte[] data) => PrepareImageBytes(data).Info;

    internal static PreparedImage PrepareImageBytes(byte[] data) {
        Guard.NotNullOrEmpty(data, nameof(data));
        if (!OfficeImageReader.TryIdentify(data, null, out OfficeImageInfo sourceInfo)) {
            // Keep the established pass-through contract for JPEG streams whose dimensions are not
            // understood by the managed header reader. The PDF writer embeds JPEG data without
            // decoding it, and layout deliberately falls back to the requested/page box in this case.
            if (LooksLikeJpeg(data)) {
                return new PreparedImage(
                    (byte[])data.Clone(),
                    new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0),
                    OfficeImageFormat.Jpeg,
                    wasTranscoded: false);
            }

            throw new NotSupportedException(SupportedImageMessage + " The source image header is not recognized.");
        }

        if (sourceInfo.Format == OfficeImageFormat.Jpeg) {
            return new PreparedImage((byte[])data.Clone(), sourceInfo, sourceInfo.Format, wasTranscoded: false);
        }

        if (sourceInfo.Format == OfficeImageFormat.Png) {
            if (PdfWriter.TryGetPngImageData(data, out _, out string? sourcePngReason)) {
                return new PreparedImage((byte[])data.Clone(), sourceInfo, sourceInfo.Format, wasTranscoded: false);
            }

            string suffix = string.IsNullOrWhiteSpace(sourcePngReason) ? string.Empty : " " + sourcePngReason;
            throw new NotSupportedException(SupportedImageMessage + suffix);
        }

        if (!OfficeImagePngConverter.TryConvertToPng(data, out byte[] normalizedPng)) {
            throw new NotSupportedException(
                $"{SupportedImageMessage} Detected {sourceInfo.Format} ({sourceInfo.MimeType}), but it could not be normalized.");
        }

        if (!PdfWriter.TryGetPngImageData(normalizedPng, out PdfWriter.PdfImageStream normalizedImage, out string? normalizedReason)) {
            string suffix = string.IsNullOrWhiteSpace(normalizedReason) ? string.Empty : " " + normalizedReason;
            throw new NotSupportedException(
                $"{SupportedImageMessage} Detected {sourceInfo.Format} ({sourceInfo.MimeType}), but it could not be normalized.{suffix}");
        }

        OfficeImageInfo normalizedInfo = OfficeImageReader.TryIdentify(
            normalizedPng,
            null,
            out OfficeImageInfo identifiedNormalized)
                ? identifiedNormalized
                : new OfficeImageInfo(
                    OfficeImageFormat.Png,
                    normalizedImage.PixelWidth,
                    normalizedImage.PixelHeight,
                    sourceInfo.DpiX,
                    sourceInfo.DpiY);
        return new PreparedImage(normalizedPng, normalizedInfo, sourceInfo.Format, wasTranscoded: true);
    }

    private static bool LooksLikeJpeg(byte[] data) =>
        data.Length >= 4 &&
        data[0] == 0xFF &&
        data[1] == 0xD8 &&
        data[data.Length - 2] == 0xFF &&
        data[data.Length - 1] == 0xD9;
}
