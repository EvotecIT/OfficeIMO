namespace OfficeIMO.Drawing {
    /// <summary>
    /// Validates image bytes against the lightweight image constraints used by first-party PDF export preflight checks.
    /// </summary>
    public static class OfficeImagePdfCompatibility {
        private static readonly byte[] PngSignature = { 137, 80, 78, 71, 13, 10, 26, 10 };

        /// <summary>
        /// Returns true when the image bytes are JPEG or structurally valid PNG bytes suitable for first-party PDF export.
        /// </summary>
        public static bool TryValidate(byte[] bytes, out OfficeImageInfo? imageInfo, out string? unsupportedReason) {
            imageInfo = null;
            unsupportedReason = null;
            if (bytes == null || bytes.Length == 0) {
                unsupportedReason = "Image bytes are empty.";
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
                    unsupportedReason = $"Detected {detected.Format} ({detected.MimeType}); first-party PDF export supports JPEG and PNG image bytes.";
                    return false;
            }
        }

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
