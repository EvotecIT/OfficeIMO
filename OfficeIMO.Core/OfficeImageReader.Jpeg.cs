using System;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private const string JpegXmpIdentifier = "http://ns.adobe.com/xap/1.0/\0";

    private static bool TryReadJpeg(byte[] data, out OfficeImageInfo info) =>
        TryReadJpeg(data, CancellationToken.None, out info);

    private static bool TryReadJpeg(byte[] data, CancellationToken cancellationToken, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 4 || data[0] != 0xFF || data[1] != 0xD8) {
            return false;
        }

        double dpiX = 96.0;
        double dpiY = 96.0;
        int offset = 2;

        while (offset < data.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            if (data[offset] != 0xFF) {
                return false;
            }

            while (offset < data.Length && data[offset] == 0xFF) {
                offset++;
            }

            if (offset >= data.Length) {
                break;
            }

            byte marker = data[offset++];
            if (marker == 0xD9 || marker == 0xDA) {
                return false;
            }
            if (marker == 0x01) {
                continue;
            }
            if (marker == 0x00 || marker == 0xD8 || (marker >= 0xD0 && marker <= 0xD7)) {
                return false;
            }

            if (offset + 2 > data.Length) {
                return false;
            }

            int segmentLength = ReadUInt16BigEndian(data, offset);
            if (segmentLength < 2 || offset + segmentLength > data.Length) {
                return false;
            }

            int segmentStart = offset + 2;
            int segmentDataLength = segmentLength - 2;

            if (marker == 0xE0 && HasJpegSegmentPrefix(data, segmentStart, segmentDataLength, "JFIF\0")) {
                if (TryReadJfifSegment(
                    data,
                    segmentStart,
                    segmentDataLength,
                    out byte units,
                    out int xDensity,
                    out int yDensity)) {
                    if (xDensity > 0 && yDensity > 0) {
                        if (units == 1) {
                            dpiX = xDensity;
                            dpiY = yDensity;
                        } else if (units == 2) {
                            dpiX = xDensity * 2.54;
                            dpiY = yDensity * 2.54;
                        }
                    }
                }
            }

            if (IsStartOfFrame(marker)) {
                if (!TryReadJpegFrameHeader(
                    data,
                    segmentStart,
                    segmentDataLength,
                    out int width,
                    out int height)) {
                    return false;
                }
                info = new OfficeImageInfo(OfficeImageFormat.Jpeg, width, height, dpiX, dpiY);
                return true;
            }

            offset += segmentLength;
        }

        return false;
    }

    private static bool TryReadJpegFrameHeader(
        byte[] data,
        int segmentStart,
        int segmentDataLength,
        out int width,
        out int height) {
        width = 0;
        height = 0;
        if (segmentDataLength < 9) return false;

        int componentCount = data[segmentStart + 5];
        if (componentCount == 0 || segmentDataLength != 6 + (componentCount * 3)) return false;
        if (data[segmentStart] == 0) return false;

        height = ReadUInt16BigEndian(data, segmentStart + 1);
        width = ReadUInt16BigEndian(data, segmentStart + 3);
        return width > 0 && height > 0;
    }

    private static bool IsStartOfFrame(byte marker) =>
        marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;

    private static bool HasCompleteJpegPayload(byte[] data) {
        if (data.Length < 12 || data[0] != 0xFF || data[1] != 0xD8) return false;
        bool hasFrame = false;
        bool hasScan = false;
        bool currentScanHasEntropyData = false;
        bool inScan = false;
        bool seenExif = false;
        bool seenJfif = false;
        bool seenXmp = false;
        byte[][]? iccSegments = null;
        int offset = 2;
        while (offset < data.Length) {
            if (inScan && data[offset] != 0xFF) {
                currentScanHasEntropyData = true;
                offset++;
                continue;
            }
            if (data[offset] != 0xFF) return false;
            while (offset < data.Length && data[offset] == 0xFF) offset++;
            if (offset >= data.Length) return false;

            byte marker = data[offset++];
            if (inScan) {
                if (marker == 0x00) {
                    currentScanHasEntropyData = true;
                    continue;
                }
                if (marker >= 0xD0 && marker <= 0xD7) continue;
                if (!currentScanHasEntropyData) return false;
                inScan = false;
            }

            if (marker == 0xD9) {
                return hasFrame && hasScan && offset == data.Length && HasValidJpegIccProfile(iccSegments);
            }
            if (marker == 0x01) continue;
            if (marker == 0x00 || marker == 0xD8 || (marker >= 0xD0 && marker <= 0xD7) || offset + 2 > data.Length) {
                return false;
            }

            int segmentLength = ReadUInt16BigEndian(data, offset);
            if (segmentLength < 2 || offset + segmentLength > data.Length) return false;
            int segmentStart = offset + 2;
            int segmentDataLength = segmentLength - 2;
            if (IsStartOfFrame(marker)) {
                if (!TryReadJpegFrameHeader(data, segmentStart, segmentDataLength, out _, out _)) return false;
                hasFrame = true;
            } else if (marker == 0xE0 && HasJpegSegmentPrefix(
                data,
                segmentStart,
                segmentDataLength,
                "JFIF\0")) {
                if (seenJfif || hasFrame || !TryReadJfifSegment(
                    data,
                    segmentStart,
                    segmentDataLength,
                    out _,
                    out _,
                    out _)) return false;
                seenJfif = true;
            } else if (marker == 0xE1 && HasJpegSegmentPrefix(data, segmentStart, segmentDataLength, "Exif\0\0")) {
                if (seenExif || !OfficeTiffStructureValidator.TryValidateExif(
                    data,
                    segmentStart + 6,
                    segmentDataLength - 6)) {
                    return false;
                }
                seenExif = true;
            } else if (marker == 0xE1 && HasJpegSegmentPrefix(
                data,
                segmentStart,
                segmentDataLength,
                JpegXmpIdentifier)) {
                int packetOffset = segmentStart + JpegXmpIdentifier.Length;
                int packetLength = segmentDataLength - JpegXmpIdentifier.Length;
                if (seenXmp || !OfficeXmpPacketValidator.TryValidate(
                    data,
                    packetOffset,
                    packetLength)) {
                    return false;
                }
                seenXmp = true;
            } else if (marker == 0xE2 && HasJpegSegmentPrefix(
                data,
                segmentStart,
                segmentDataLength,
                "ICC_PROFILE\0")) {
                if (segmentDataLength <= 14) return false;
                int sequence = data[segmentStart + 12];
                int segmentCount = data[segmentStart + 13];
                if (sequence <= 0 || segmentCount <= 0 || sequence > segmentCount) return false;
                if (iccSegments == null) {
                    iccSegments = new byte[segmentCount][];
                } else if (iccSegments.Length != segmentCount) {
                    return false;
                }
                if (iccSegments[sequence - 1] != null) return false;
                int profilePartLength = segmentDataLength - 14;
                var profilePart = new byte[profilePartLength];
                Buffer.BlockCopy(data, segmentStart + 14, profilePart, 0, profilePartLength);
                iccSegments[sequence - 1] = profilePart;
            } else if (marker == 0xDA) {
                if (!hasFrame || segmentDataLength < 6) return false;
                int componentCount = data[segmentStart];
                if (componentCount == 0 || segmentDataLength != 4 + (componentCount * 2)) return false;
                hasScan = true;
                currentScanHasEntropyData = false;
                inScan = true;
            }

            offset += segmentLength;
        }
        return false;
    }

    private static bool HasJpegSegmentPrefix(
        byte[] data,
        int offset,
        int count,
        string prefix) {
        if (count < prefix.Length) return false;
        for (int index = 0; index < prefix.Length; index++) {
            if (data[offset + index] != (byte)prefix[index]) return false;
        }
        return true;
    }

    private static bool TryReadJfifSegment(
        byte[] data,
        int offset,
        int length,
        out byte units,
        out int xDensity,
        out int yDensity) {
        units = 0;
        xDensity = 0;
        yDensity = 0;
        if (length < 14 || data[offset + 5] != 1 || data[offset + 6] > 2) return false;
        units = data[offset + 7];
        xDensity = ReadUInt16BigEndian(data, offset + 8);
        yDensity = ReadUInt16BigEndian(data, offset + 10);
        int thumbnailWidth = data[offset + 12];
        int thumbnailHeight = data[offset + 13];
        long expectedLength = 14L + 3L * thumbnailWidth * thumbnailHeight;
        return units <= 2 && xDensity > 0 && yDensity > 0 && expectedLength == length;
    }

    private static bool HasValidJpegIccProfile(byte[][]? segments) {
        if (segments == null) return true;
        int profileLength = 0;
        for (int index = 0; index < segments.Length; index++) {
            byte[]? segment = segments[index];
            if (segment == null || segment.Length > OfficeRasterGuards.MaximumEncodedBytes - profileLength) {
                return false;
            }
            profileLength += segment.Length;
        }
        var profile = new byte[profileLength];
        int offset = 0;
        for (int index = 0; index < segments.Length; index++) {
            byte[] segment = segments[index];
            Buffer.BlockCopy(segment, 0, profile, offset, segment.Length);
            offset += segment.Length;
        }
        return OfficeIccProfileValidator.TryValidate(profile, 0, profile.Length);
    }
}
