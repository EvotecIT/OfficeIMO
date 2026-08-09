using System;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>Validates PNG chunk framing, ordering, and CRC integrity without decoding pixels.</summary>
internal static class OfficePngContainerValidator {
    private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };

    internal static bool TryValidate(byte[]? bytes, out int frameCount, out string? failureReason) {
        frameCount = 0;
        failureReason = null;
        if (bytes == null || bytes.Length < 33 || !HasSignature(bytes)) {
            failureReason = "PNG bytes are missing the PNG signature or required chunks.";
            return false;
        }

        try {
            OfficeRasterGuards.EnsurePayloadWithinLimits(bytes.Length, "PNG payload exceeds size limits.");
            bool seenHeader = false;
            bool seenImageData = false;
            bool seenAnimationControl = false;
            int declaredFrameCount = 1;
            int offset = Signature.Length;
            while (offset + 12 <= bytes.Length) {
                int length = ReadBigEndianInt32(bytes, offset);
                long chunkEnd = (long)offset + 12L + length;
                if (length < 0 || chunkEnd > bytes.Length) {
                    failureReason = "PNG chunk length exceeds the available image bytes.";
                    return false;
                }

                string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
                if (!seenHeader && (type != "IHDR" || length != 13)) {
                    failureReason = "PNG bytes must start with an IHDR chunk.";
                    return false;
                }

                uint expectedCrc = ReadBigEndianUInt32(bytes, offset + 8 + length);
                uint actualCrc = ComputeCrc(bytes, offset + 4, 4 + length);
                if (actualCrc != expectedCrc) {
                    failureReason = "PNG chunk '" + type + "' has an invalid CRC.";
                    return false;
                }

                int dataOffset = offset + 8;
                switch (type) {
                    case "IHDR":
                        if (seenHeader || length != 13) {
                            failureReason = "PNG bytes contain an invalid or repeated IHDR chunk.";
                            return false;
                        }
                        seenHeader = true;
                        break;
                    case "acTL":
                        if (!seenHeader || seenImageData || seenAnimationControl || length != 8) {
                            failureReason = "PNG bytes contain an invalid APNG animation-control chunk.";
                            return false;
                        }
                        int candidate = ReadBigEndianInt32(bytes, dataOffset);
                        if (candidate <= 0) {
                            failureReason = "PNG animation frame count must be positive.";
                            return false;
                        }
                        declaredFrameCount = candidate;
                        seenAnimationControl = true;
                        break;
                    case "IDAT":
                        if (!seenHeader) {
                            failureReason = "PNG image data appears before the header.";
                            return false;
                        }
                        seenImageData = true;
                        break;
                    case "IEND":
                        if (length != 0 || !seenImageData) {
                            failureReason = "PNG bytes contain an invalid IEND chunk or no image data.";
                            return false;
                        }
                        offset = (int)chunkEnd;
                        if (offset != bytes.Length) {
                            failureReason = "PNG bytes contain trailing data after IEND.";
                            return false;
                        }
                        frameCount = declaredFrameCount;
                        return true;
                }

                offset = (int)chunkEnd;
            }
        } catch (Exception exception) when (exception is FormatException || exception is OverflowException) {
            failureReason = exception.Message;
            return false;
        }

        failureReason = "PNG bytes do not contain a complete IEND chunk.";
        return false;
    }

    private static bool HasSignature(byte[] bytes) {
        for (int index = 0; index < Signature.Length; index++) {
            if (bytes[index] != Signature[index]) return false;
        }
        return true;
    }

    private static uint ComputeCrc(byte[] bytes, int offset, int count) {
        uint crc = 0xFFFFFFFFU;
        for (int index = 0; index < count; index++) {
            crc ^= bytes[offset + index];
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1U) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
            }
        }
        return crc ^ 0xFFFFFFFFU;
    }

    private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
        (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    private static uint ReadBigEndianUInt32(byte[] bytes, int offset) =>
        ((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) | ((uint)bytes[offset + 2] << 8) | bytes[offset + 3];
}
