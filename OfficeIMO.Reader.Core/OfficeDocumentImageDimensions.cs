using System;

namespace OfficeIMO.Reader;

/// <summary>
/// Reads pixel dimensions from common encoded image formats without decoding pixels.
/// </summary>
public static class OfficeDocumentImageDimensions {
    /// <summary>
    /// Attempts to read encoded image dimensions from PNG, JPEG, GIF, or BMP bytes.
    /// </summary>
    public static bool TryReadPixelDimensions(byte[] payload, string? mediaType, out int width, out int height) {
        width = 0;
        height = 0;
        if (payload == null || payload.Length < 10) {
            return false;
        }

        string normalizedMediaType = mediaType?.Trim().ToLowerInvariant() ?? string.Empty;
        if (normalizedMediaType == "image/png" || HasPngSignature(payload)) {
            return TryReadPng(payload, out width, out height);
        }

        if (normalizedMediaType == "image/gif" || HasGifSignature(payload)) {
            return TryReadGif(payload, out width, out height);
        }

        if (normalizedMediaType == "image/bmp" || HasBmpSignature(payload)) {
            return TryReadBmp(payload, out width, out height);
        }

        if (normalizedMediaType == "image/jpeg" ||
            normalizedMediaType == "image/jpg" ||
            HasJpegSignature(payload)) {
            return TryReadJpeg(payload, out width, out height);
        }

        return false;
    }

    private static bool TryReadPng(byte[] payload, out int width, out int height) {
        width = 0;
        height = 0;
        if (payload.Length < 24 || !HasPngSignature(payload)) {
            return false;
        }

        width = ReadInt32BigEndian(payload, 16);
        height = ReadInt32BigEndian(payload, 20);
        return width > 0 && height > 0;
    }

    private static bool TryReadGif(byte[] payload, out int width, out int height) {
        width = 0;
        height = 0;
        if (payload.Length < 10 || !HasGifSignature(payload)) {
            return false;
        }

        width = ReadUInt16LittleEndian(payload, 6);
        height = ReadUInt16LittleEndian(payload, 8);
        return width > 0 && height > 0;
    }

    private static bool TryReadBmp(byte[] payload, out int width, out int height) {
        width = 0;
        height = 0;
        if (payload.Length < 26 || !HasBmpSignature(payload)) {
            return false;
        }

        width = ReadInt32LittleEndian(payload, 18);
        height = Math.Abs(ReadInt32LittleEndian(payload, 22));
        return width > 0 && height > 0;
    }

    private static bool TryReadJpeg(byte[] payload, out int width, out int height) {
        width = 0;
        height = 0;
        if (payload.Length < 4 || !HasJpegSignature(payload)) {
            return false;
        }

        int index = 2;
        while (index + 3 < payload.Length) {
            if (payload[index] != 0xFF) {
                index++;
                continue;
            }

            while (index < payload.Length && payload[index] == 0xFF) {
                index++;
            }

            if (index >= payload.Length) {
                return false;
            }

            byte marker = payload[index++];
            if (marker == 0xD9 || marker == 0xDA) {
                return false;
            }

            if (index + 1 >= payload.Length) {
                return false;
            }

            int segmentLength = ReadUInt16BigEndian(payload, index);
            if (segmentLength < 2 || index + segmentLength > payload.Length) {
                return false;
            }

            if (IsJpegStartOfFrame(marker)) {
                if (segmentLength < 7) {
                    return false;
                }

                height = ReadUInt16BigEndian(payload, index + 3);
                width = ReadUInt16BigEndian(payload, index + 5);
                return width > 0 && height > 0;
            }

            index += segmentLength;
        }

        return false;
    }

    private static bool IsJpegStartOfFrame(byte marker) {
        return marker == 0xC0 ||
               marker == 0xC1 ||
               marker == 0xC2 ||
               marker == 0xC3 ||
               marker == 0xC5 ||
               marker == 0xC6 ||
               marker == 0xC7 ||
               marker == 0xC9 ||
               marker == 0xCA ||
               marker == 0xCB ||
               marker == 0xCD ||
               marker == 0xCE ||
               marker == 0xCF;
    }

    private static bool HasPngSignature(byte[] payload) {
        return payload.Length >= 8 &&
               payload[0] == 0x89 &&
               payload[1] == 0x50 &&
               payload[2] == 0x4E &&
               payload[3] == 0x47 &&
               payload[4] == 0x0D &&
               payload[5] == 0x0A &&
               payload[6] == 0x1A &&
               payload[7] == 0x0A;
    }

    private static bool HasGifSignature(byte[] payload) {
        return payload.Length >= 6 &&
               payload[0] == 0x47 &&
               payload[1] == 0x49 &&
               payload[2] == 0x46 &&
               payload[3] == 0x38 &&
               (payload[4] == 0x37 || payload[4] == 0x39) &&
               payload[5] == 0x61;
    }

    private static bool HasBmpSignature(byte[] payload) {
        return payload.Length >= 2 && payload[0] == 0x42 && payload[1] == 0x4D;
    }

    private static bool HasJpegSignature(byte[] payload) {
        return payload.Length >= 2 && payload[0] == 0xFF && payload[1] == 0xD8;
    }

    private static int ReadUInt16BigEndian(byte[] payload, int offset) {
        return (payload[offset] << 8) | payload[offset + 1];
    }

    private static int ReadUInt16LittleEndian(byte[] payload, int offset) {
        return payload[offset] | (payload[offset + 1] << 8);
    }

    private static int ReadInt32BigEndian(byte[] payload, int offset) {
        return (payload[offset] << 24) |
               (payload[offset + 1] << 16) |
               (payload[offset + 2] << 8) |
               payload[offset + 3];
    }

    private static int ReadInt32LittleEndian(byte[] payload, int offset) {
        return payload[offset] |
               (payload[offset + 1] << 8) |
               (payload[offset + 2] << 16) |
               (payload[offset + 3] << 24);
    }
}
