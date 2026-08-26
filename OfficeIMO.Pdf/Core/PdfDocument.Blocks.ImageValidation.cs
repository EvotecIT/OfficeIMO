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
                OfficeImageMetadataSnapshot jpegMetadata = OfficeImageMetadataInspector.Inspect(
                    data,
                    OfficeImageFormat.Jpeg);
                bool hasJpegIcc = (jpegMetadata.Kinds & OfficeImageMetadataKinds.Icc) != 0;
                if (hasJpegIcc && jpegMetadata.Icc == null) {
                    throw new NotSupportedException(
                        SupportedImageMessage + " The embedded JPEG ICC profile cannot be retained or normalized safely.");
                }
                if (hasJpegIcc) {
                    if (!PdfWriter.TryGetJpegComponentCount(data, out int jpegComponentCount)) {
                        throw new NotSupportedException(
                            SupportedImageMessage + " The tagged JPEG component count cannot be verified; four-component JPEG data cannot be normalized safely.");
                    }
                    if (jpegComponentCount == 4) {
                        throw new NotSupportedException(
                            SupportedImageMessage + " A four-component JPEG with an embedded ICC profile cannot be normalized safely.");
                    }
                }
                return new PreparedImage(
                    (byte[])data.Clone(),
                    new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0),
                    OfficeImageFormat.Jpeg,
                    wasTranscoded: false);
            }

            // Metadata identification deliberately rejects PNG dimensions that exceed the
            // shared raster budget. Preserve the PDF writer's more specific validation
            // diagnostic instead of collapsing those payloads into an unknown header.
            if (LooksLikePng(data)) {
                if (PdfWriter.TryGetPngImageData(data, out PdfWriter.PdfImageStream pngImage, out string? pngReason)) {
                    return new PreparedImage(
                        (byte[])data.Clone(),
                        new OfficeImageInfo(OfficeImageFormat.Png, pngImage.PixelWidth, pngImage.PixelHeight),
                        OfficeImageFormat.Png,
                        wasTranscoded: false);
                }

                string suffix = string.IsNullOrWhiteSpace(pngReason) ? string.Empty : " " + pngReason;
                throw new NotSupportedException(SupportedImageMessage + suffix);
            }

            throw new NotSupportedException(SupportedImageMessage + " The source image header is not recognized.");
        }

        OfficeImageMetadataSnapshot sourceMetadata = OfficeImageMetadataInspector.Inspect(data, sourceInfo.Format);
        bool hasEmbeddedIccProfile = (sourceMetadata.Kinds & OfficeImageMetadataKinds.Icc) != 0;
        if (hasEmbeddedIccProfile && sourceMetadata.Icc == null) {
            throw new NotSupportedException(
                SupportedImageMessage + " The embedded ICC profile cannot be retained or normalized safely.");
        }

        if (sourceInfo.Format == OfficeImageFormat.Jpeg) {
            bool hasComponentCount = PdfWriter.TryGetJpegComponentCount(data, out int componentCount);
            if (hasEmbeddedIccProfile && !hasComponentCount) {
                throw new NotSupportedException(
                    SupportedImageMessage + " The tagged JPEG component count cannot be verified; four-component JPEG data cannot be normalized safely.");
            }
            if (hasComponentCount && componentCount == 4) {
                if (hasEmbeddedIccProfile) {
                    throw new NotSupportedException(
                        SupportedImageMessage + " A four-component JPEG with an embedded ICC profile cannot be normalized safely.");
                }
                if (!OfficeImagePdfCompatibility.TryValidateTranscodeDimensions(
                        sourceInfo,
                        OfficeImagePdfCompatibility.DefaultMaximumTranscodePixels,
                        out string? jpegTranscodeLimitReason)) {
                    throw new NotSupportedException(SupportedImageMessage + " " + jpegTranscodeLimitReason);
                }
                if (!OfficeImagePngConverter.TryConvertToPng(data, out byte[] normalizedJpegPng) ||
                    !OfficeImageReader.TryIdentify(normalizedJpegPng, null, out OfficeImageInfo normalizedJpegInfo)) {
                    throw new NotSupportedException(SupportedImageMessage + " Four-component JPEG data could not be normalized safely for PDF embedding.");
                }
                return new PreparedImage(normalizedJpegPng, normalizedJpegInfo, sourceInfo.Format, wasTranscoded: true);
            }
            if (componentCount != 0 && componentCount != 1 && componentCount != 3) {
                throw new NotSupportedException(SupportedImageMessage + " JPEG component count is not supported for PDF embedding.");
            }
            return new PreparedImage((byte[])data.Clone(), sourceInfo, sourceInfo.Format, wasTranscoded: false);
        }

        if (sourceInfo.Format == OfficeImageFormat.Png) {
            if (PdfWriter.TryGetPngImageData(data, out _, out string? sourcePngReason)) {
                return new PreparedImage((byte[])data.Clone(), sourceInfo, sourceInfo.Format, wasTranscoded: false);
            }

            string suffix = string.IsNullOrWhiteSpace(sourcePngReason) ? string.Empty : " " + sourcePngReason;
            throw new NotSupportedException(SupportedImageMessage + suffix);
        }

        if (hasEmbeddedIccProfile) {
            throw new NotSupportedException(
                SupportedImageMessage + $" The embedded {sourceInfo.Format} ICC profile cannot be retained through PNG normalization.");
        }

        if (!OfficeImagePdfCompatibility.TryValidateTranscodeDimensions(
                sourceInfo,
                OfficeImagePdfCompatibility.DefaultMaximumTranscodePixels,
                out string? transcodeLimitReason)) {
            throw new NotSupportedException(SupportedImageMessage + " " + transcodeLimitReason);
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

    private static bool LooksLikePng(byte[] data) =>
        data.Length >= 8 &&
        data[0] == 137 && data[1] == 80 && data[2] == 78 && data[3] == 71 &&
        data[4] == 13 && data[5] == 10 && data[6] == 26 && data[7] == 10;
}
