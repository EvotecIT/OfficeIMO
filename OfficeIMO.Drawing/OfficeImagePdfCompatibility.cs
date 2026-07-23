namespace OfficeIMO.Drawing {
    /// <summary>
    /// Validates image bytes against the shared Drawing raster contract used by first-party PDF export preflight checks.
    /// </summary>
    public static class OfficeImagePdfCompatibility {
        /// <summary>Default pixel ceiling for raster formats that must be decoded and re-encoded before PDF embedding.</summary>
        public const long DefaultMaximumTranscodePixels = 8_000_000L;

        private static readonly byte[] PngSignature = { 137, 80, 78, 71, 13, 10, 26, 10 };

        /// <summary>
        /// Returns true when the image bytes can be embedded directly or normalized by the shared Drawing raster engine.
        /// </summary>
        public static bool TryValidate(byte[] bytes, out OfficeImageInfo? imageInfo, out string? unsupportedReason) {
            return TryValidate(bytes, DefaultMaximumTranscodePixels, out imageInfo, out unsupportedReason);
        }

        /// <summary>
        /// Returns true when the image bytes can be embedded directly or normalized without exceeding the supplied transcode pixel budget.
        /// </summary>
        public static bool TryValidate(byte[] bytes, long maximumTranscodePixels, out OfficeImageInfo? imageInfo, out string? unsupportedReason) {
            imageInfo = null;
            unsupportedReason = null;
            if (maximumTranscodePixels < 1) throw new System.ArgumentOutOfRangeException(nameof(maximumTranscodePixels));
            if (bytes == null || bytes.Length == 0) {
                unsupportedReason = "Image bytes are empty.";
                return false;
            }

            // Preserve actionable PNG diagnostics even when the general image reader
            // correctly rejects a malformed container during format identification.
            if (StartsWith(bytes, PngSignature) && !TryValidatePngContainer(bytes, out unsupportedReason)) {
                return false;
            }

            if (!OfficeImageReader.TryIdentify(bytes, null, out OfficeImageInfo detected)) {
                unsupportedReason = "Image bytes do not contain a supported image header.";
                return false;
            }

            imageInfo = detected;
            switch (detected.Format) {
                case OfficeImageFormat.Jpeg:
                    return true;
                case OfficeImageFormat.Png:
                    return TryValidatePngContainer(bytes, out unsupportedReason);
                default:
                    if (!IsSupportedFormat(detected.Format)) {
                        unsupportedReason = $"Detected {detected.Format} ({detected.MimeType}); the format is not part of the shared PDF raster contract.";
                        return false;
                    }

                    if (!TryValidateTranscodeDimensions(detected, maximumTranscodePixels, out unsupportedReason)) return false;

                    if (!OfficeImagePngConverter.TryConvertToPng(bytes, out byte[] normalizedPng) ||
                        !TryValidatePngContainer(normalizedPng, out unsupportedReason)) {
                        unsupportedReason ??= $"Detected {detected.Format} ({detected.MimeType}), but the payload could not be normalized by OfficeIMO.Drawing.";
                        return false;
                    }

                    return true;
            }
        }

        /// <summary>
        /// Validates identified source dimensions before a PDF path decodes and re-encodes a raster payload.
        /// </summary>
        public static bool TryValidateTranscodeDimensions(
            OfficeImageInfo imageInfo,
            long maximumTranscodePixels,
            out string? unsupportedReason) {
            if (imageInfo == null) throw new System.ArgumentNullException(nameof(imageInfo));
            if (maximumTranscodePixels < 1) throw new System.ArgumentOutOfRangeException(nameof(maximumTranscodePixels));

            unsupportedReason = null;
            long pixels = checked((long)imageInfo.Width * imageInfo.Height);
            if (imageInfo.Width < 1 || imageInfo.Height < 1 || pixels > maximumTranscodePixels) {
                unsupportedReason = $"Detected {imageInfo.Format} requires raster transcoding of {pixels} pixels, exceeding the configured limit of {maximumTranscodePixels} pixels.";
                return false;
            }

            return true;
        }

        /// <summary>
        /// Returns true when the MIME content type belongs to the shared PDF raster source set.
        /// </summary>
        public static bool IsSupportedContentType(string? contentType) =>
            TryGetSupportedContentTypeFormat(contentType, out _);

        /// <summary>
        /// Resolves a MIME content type accepted by the shared PDF raster source set to its image format.
        /// </summary>
        public static bool TryGetSupportedContentTypeFormat(string? contentType, out OfficeImageFormat format) {
            format = OfficeImageInfo.FromMimeType(contentType);
            return IsSupportedFormat(format);
        }

        /// <summary>
        /// Validates that image bytes can be identified and match the declared PDF-supported MIME content type.
        /// </summary>
        /// <remarks>
        /// This checks shared image identity and declared MIME parity only. Format-specific PDF writer validation
        /// remains owned by the PDF writer because it depends on PDF stream encoding constraints.
        /// </remarks>
        /// <param name="bytes">Source image bytes.</param>
        /// <param name="contentType">Declared MIME content type.</param>
        /// <param name="imageInfo">Detected image metadata when the bytes can be identified.</param>
        /// <param name="unsupportedReason">Reason the declared content type and detected bytes are not compatible.</param>
        /// <returns><c>true</c> when the declared PDF image format matches the detected image bytes.</returns>
        public static bool TryValidateDeclaredContentType(byte[] bytes, string? contentType, out OfficeImageInfo? imageInfo, out string? unsupportedReason) {
            imageInfo = null;
            unsupportedReason = null;
            if (bytes == null || bytes.Length == 0) {
                unsupportedReason = "Image bytes are empty.";
                return false;
            }

            if (!TryGetSupportedContentTypeFormat(contentType, out OfficeImageFormat declaredFormat)) {
                unsupportedReason = $"Image content type '{contentType}' is not supported by first-party PDF export.";
                return false;
            }

            if (!OfficeImageReader.TryIdentify(bytes, null, out OfficeImageInfo detected)) {
                unsupportedReason = "Image bytes do not contain a supported image header.";
                return false;
            }

            imageInfo = detected;
            if (detected.Format != declaredFormat) {
                unsupportedReason = $"Image bytes were declared as {GetPdfImageFormatDisplayName(declaredFormat)} but were detected as {detected.Format}.";
                return false;
            }

            return true;
        }

        /// <summary>
        /// Returns true when the shared image format can be embedded directly or normalized by OfficeIMO.Drawing.
        /// </summary>
        public static bool IsSupportedFormat(OfficeImageFormat format) =>
            format == OfficeImageFormat.Png ||
            format == OfficeImageFormat.Jpeg ||
            format == OfficeImageFormat.Gif ||
            format == OfficeImageFormat.Bmp ||
            format == OfficeImageFormat.Tiff ||
            format == OfficeImageFormat.Webp;

        private static string GetPdfImageFormatDisplayName(OfficeImageFormat format) =>
            format == OfficeImageFormat.Jpeg ? "JPEG" : format.ToString().ToUpperInvariant();

        private static bool TryValidatePngContainer(byte[] bytes, out string? unsupportedReason) {
            unsupportedReason = null;
            if (bytes.Length < 33 || !StartsWith(bytes, PngSignature)) {
                unsupportedReason = "PNG bytes are missing the PNG signature or required chunks.";
                return false;
            }

            int offset = 8;
            bool seenIhdr = false;
            bool seenIdat = false;
            while (offset + 12 <= bytes.Length) {
                int length = ReadInt32BigEndian(bytes, offset);
                long chunkEnd = (long)offset + 12L + length;
                if (length < 0 || chunkEnd > bytes.Length) {
                    unsupportedReason = "PNG chunk length exceeds the available image bytes.";
                    return false;
                }

                string type = GetAscii(bytes, offset + 4, 4);
                uint expectedCrc = ReadUInt32BigEndian(bytes, offset + 8 + length);
                uint actualCrc = ComputePngCrc(bytes, offset + 4, 4 + length);
                if (actualCrc != expectedCrc) {
                    unsupportedReason = $"PNG chunk '{type}' has an invalid CRC.";
                    return false;
                }

                if (!seenIhdr) {
                    if (type != "IHDR" || length != 13) {
                        unsupportedReason = "PNG bytes must start with an IHDR chunk.";
                        return false;
                    }

                    seenIhdr = true;
                } else if (type == "IHDR") {
                    unsupportedReason = "PNG bytes contain more than one IHDR chunk.";
                    return false;
                }

                if (type == "IDAT") {
                    seenIdat = true;
                }

                offset = (int)chunkEnd;
                if (type == "IEND") {
                    if (!seenIdat) {
                        unsupportedReason = "PNG bytes do not contain image data.";
                        return false;
                    }

                    return offset == bytes.Length;
                }
            }

            unsupportedReason = "PNG bytes do not contain a complete IEND chunk.";
            return false;
        }

        private static bool StartsWith(byte[] data, byte[] prefix) {
            if (data.Length < prefix.Length) {
                return false;
            }

            for (int i = 0; i < prefix.Length; i++) {
                if (data[i] != prefix[i]) {
                    return false;
                }
            }

            return true;
        }

        private static int ReadInt32BigEndian(byte[] data, int offset) {
            return (data[offset] << 24)
                   | (data[offset + 1] << 16)
                   | (data[offset + 2] << 8)
                   | data[offset + 3];
        }

        private static uint ReadUInt32BigEndian(byte[] data, int offset) {
            return ((uint)data[offset] << 24)
                   | ((uint)data[offset + 1] << 16)
                   | ((uint)data[offset + 2] << 8)
                   | data[offset + 3];
        }

        private static string GetAscii(byte[] data, int offset, int count) {
            char[] chars = new char[count];
            for (int i = 0; i < count; i++) {
                chars[i] = (char)data[offset + i];
            }

            return new string(chars);
        }

        private static uint ComputePngCrc(byte[] bytes, int offset, int count) {
            uint crc = 0xFFFFFFFFU;
            for (int i = 0; i < count; i++) {
                crc = UpdateCrc(crc, bytes[offset + i]);
            }

            return crc ^ 0xFFFFFFFFU;
        }

        private static uint UpdateCrc(uint crc, byte value) {
            crc ^= value;
            for (int i = 0; i < 8; i++) {
                crc = (crc & 1U) != 0
                    ? 0xEDB88320U ^ (crc >> 1)
                    : crc >> 1;
            }

            return crc;
        }
    }
}
