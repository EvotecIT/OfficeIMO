using System;
using System.Globalization;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Header-only image metadata reader used to avoid full image decoding dependencies.
/// </summary>
public static class OfficeImageReader {
    /// <summary>
    /// Identifies image metadata from a file.
    /// </summary>
    public static OfficeImageInfo Identify(string filePath) {
        using var stream = File.OpenRead(filePath);
        if (TryIdentify(stream, filePath, out var info)) {
            return info;
        }

        throw new NotSupportedException($"Image format is not supported: {filePath}");
    }

    /// <summary>
    /// Identifies image metadata from a byte array.
    /// </summary>
    public static OfficeImageInfo Identify(byte[] data, string? fileName = null) {
        if (TryIdentify(data, fileName, out var info)) {
            return info;
        }

        throw new NotSupportedException("Image format is not supported.");
    }

    /// <summary>
    /// Tries to identify image metadata from a stream.
    /// </summary>
    public static bool TryIdentify(Stream stream, string? fileName, out OfficeImageInfo info) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        var originalPosition = stream.CanSeek ? stream.Position : 0;
        try {
            byte[] data;
            using (var ms = new MemoryStream()) {
                stream.CopyTo(ms);
                data = ms.ToArray();
            }

            return TryIdentify(data, fileName, out info);
        } finally {
            if (stream.CanSeek) {
                stream.Position = originalPosition;
            }
        }
    }

    /// <summary>
    /// Tries to identify image metadata from a byte array.
    /// </summary>
    public static bool TryIdentify(byte[]? data, string? fileName, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data == null || data.Length == 0) {
            return false;
        }

        if (TryReadPng(data, out info) ||
            TryReadJpeg(data, out info) ||
            TryReadGif(data, out info) ||
            TryReadBmp(data, out info) ||
            TryReadWebp(data, out info) ||
            TryReadTiff(data, out info) ||
            TryReadIcon(data, out info) ||
            TryReadPcx(data, out info) ||
            TryReadEmf(data, out info) ||
            TryReadWmf(data, out info) ||
            TryReadWebp(data, out info) ||
            TryReadSvg(data, fileName, out info)) {
            return true;
        }

        var byExtension = FromExtension(fileName);
        if (byExtension != OfficeImageFormat.Unknown) {
            info = new OfficeImageInfo(byExtension, 0, 0);
            return true;
        }

        return false;
    }

    /// <summary>
    /// Maps a file name or extension to a supported image format.
    /// </summary>
    public static OfficeImageFormat FromExtension(string? fileName) {
        if (string.IsNullOrWhiteSpace(fileName)) {
            return OfficeImageFormat.Unknown;
        }

        var text = fileName!.Trim();
        var hasDirectorySeparator = text.IndexOfAny(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }) >= 0;
        var ext = text.StartsWith(".", StringComparison.Ordinal) && !hasDirectorySeparator
            ? text
            : Path.GetExtension(text);
        if (string.IsNullOrEmpty(ext) && !hasDirectorySeparator && text.IndexOf('.') < 0) {
            ext = "." + text;
        }
        ext = ext.ToLowerInvariant();
        return ext switch {
            ".png" => OfficeImageFormat.Png,
            ".jpg" or ".jpeg" => OfficeImageFormat.Jpeg,
            ".gif" => OfficeImageFormat.Gif,
            ".bmp" => OfficeImageFormat.Bmp,
            ".tif" or ".tiff" => OfficeImageFormat.Tiff,
            ".svg" => OfficeImageFormat.Svg,
            ".emf" => OfficeImageFormat.Emf,
            ".wmf" => OfficeImageFormat.Wmf,
            ".ico" => OfficeImageFormat.Icon,
            ".pcx" => OfficeImageFormat.Pcx,
            ".webp" => OfficeImageFormat.Webp,
            _ => OfficeImageFormat.Unknown
        };
    }

    /// <summary>
    /// Returns whether the file name or extension maps to an image format known by the shared drawing layer.
    /// </summary>
    /// <param name="fileName">File name, path, or bare extension.</param>
    /// <returns><c>true</c> when the extension maps to a known image format.</returns>
    public static bool IsKnownImageExtension(string? fileName) => FromExtension(fileName) != OfficeImageFormat.Unknown;

    private static bool TryReadPng(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
        if (data.Length < 24 || !StartsWith(data, signature)) {
            return false;
        }

        int width = ReadInt32BigEndian(data, 16);
        int height = ReadInt32BigEndian(data, 20);
        double dpiX = 96.0;
        double dpiY = 96.0;

        int offset = 8;
        while (offset + 12 <= data.Length) {
            int length = ReadInt32BigEndian(data, offset);
            long chunkEnd = (long)offset + 12L + length;
            if (length < 0 || chunkEnd > data.Length) {
                break;
            }

            string type = GetAscii(data, offset + 4, 4);
            if (type == "pHYs" && length >= 9) {
                int xPpm = ReadInt32BigEndian(data, offset + 8);
                int yPpm = ReadInt32BigEndian(data, offset + 12);
                byte unit = data[offset + 16];
                if (unit == 1 && xPpm > 0 && yPpm > 0) {
                    dpiX = xPpm * 0.0254;
                    dpiY = yPpm * 0.0254;
                }

                break;
            }

            offset = (int)chunkEnd;
        }

        info = new OfficeImageInfo(OfficeImageFormat.Png, width, height, dpiX, dpiY);
        return width > 0 && height > 0;
    }

    private static bool TryReadJpeg(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 4 || data[0] != 0xFF || data[1] != 0xD8) {
            return false;
        }

        double dpiX = 96.0;
        double dpiY = 96.0;
        int offset = 2;

        while (offset + 4 <= data.Length) {
            if (data[offset] != 0xFF) {
                offset++;
                continue;
            }

            while (offset < data.Length && data[offset] == 0xFF) {
                offset++;
            }

            if (offset >= data.Length) {
                break;
            }

            byte marker = data[offset++];
            if (marker == 0xD9 || marker == 0xDA) {
                break;
            }

            if (offset + 2 > data.Length) {
                break;
            }

            int segmentLength = ReadUInt16BigEndian(data, offset);
            if (segmentLength < 2 || offset + segmentLength > data.Length) {
                break;
            }

            int segmentStart = offset + 2;
            int segmentDataLength = segmentLength - 2;

            if (marker == 0xE0 && segmentDataLength >= 12 && GetAscii(data, segmentStart, 5) == "JFIF\0") {
                byte units = data[segmentStart + 7];
                int xDensity = ReadUInt16BigEndian(data, segmentStart + 8);
                int yDensity = ReadUInt16BigEndian(data, segmentStart + 10);
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

            if (IsStartOfFrame(marker) && segmentDataLength >= 7) {
                int height = ReadUInt16BigEndian(data, segmentStart + 1);
                int width = ReadUInt16BigEndian(data, segmentStart + 3);
                info = new OfficeImageInfo(OfficeImageFormat.Jpeg, width, height, dpiX, dpiY);
                return width > 0 && height > 0;
            }

            offset += segmentLength;
        }

        info = new OfficeImageInfo(OfficeImageFormat.Jpeg, 0, 0, dpiX, dpiY);
        return true;
    }

    private static bool TryReadGif(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 10) {
            return false;
        }

        string signature = GetAscii(data, 0, 6);
        if (signature != "GIF87a" && signature != "GIF89a") {
            return false;
        }

        int width = ReadUInt16LittleEndian(data, 6);
        int height = ReadUInt16LittleEndian(data, 8);
        info = new OfficeImageInfo(OfficeImageFormat.Gif, width, height);
        return width > 0 && height > 0;
    }

    private static bool TryReadBmp(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 26 || data[0] != (byte)'B' || data[1] != (byte)'M') {
            return false;
        }

        int dibSize = ReadInt32LittleEndian(data, 14);
        if (dibSize < 12) {
            return false;
        }

        int width;
        int height;
        double dpiX = 96.0;
        double dpiY = 96.0;

        if (dibSize == 12 && data.Length >= 26) {
            width = ReadUInt16LittleEndian(data, 18);
            height = ReadUInt16LittleEndian(data, 20);
        } else if (data.Length >= 42) {
            width = ReadInt32LittleEndian(data, 18);
            height = Math.Abs(ReadInt32LittleEndian(data, 22));
            int xPpm = ReadInt32LittleEndian(data, 38);
            int yPpm = ReadInt32LittleEndian(data, 42);
            if (xPpm > 0) dpiX = xPpm * 0.0254;
            if (yPpm > 0) dpiY = yPpm * 0.0254;
        } else {
            return false;
        }

        info = new OfficeImageInfo(OfficeImageFormat.Bmp, width, height, dpiX, dpiY);
        return width > 0 && height > 0;
    }

    private static bool TryReadTiff(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 8) {
            return false;
        }

        bool littleEndian;
        if (data[0] == (byte)'I' && data[1] == (byte)'I') {
            littleEndian = true;
        } else if (data[0] == (byte)'M' && data[1] == (byte)'M') {
            littleEndian = false;
        } else {
            return false;
        }

        int magic = ReadUInt16(data, 2, littleEndian);
        if (magic != 42) {
            return false;
        }

        int ifdOffset = ReadInt32(data, 4, littleEndian);
        if (ifdOffset < 0 || ifdOffset + 2 > data.Length) {
            return false;
        }

        int entryCount = ReadUInt16(data, ifdOffset, littleEndian);
        int width = 0;
        int height = 0;
        double dpiX = 96.0;
        double dpiY = 96.0;
        int unit = 2;

        for (int i = 0; i < entryCount; i++) {
            int entry = ifdOffset + 2 + (i * 12);
            if (entry + 12 > data.Length) {
                break;
            }

            int tag = ReadUInt16(data, entry, littleEndian);
            int type = ReadUInt16(data, entry + 2, littleEndian);
            int count = ReadInt32(data, entry + 4, littleEndian);
            int valueOrOffset = ReadInt32(data, entry + 8, littleEndian);

            if (tag == 256) width = ReadTiffScalar(data, type, count, valueOrOffset, littleEndian);
            else if (tag == 257) height = ReadTiffScalar(data, type, count, valueOrOffset, littleEndian);
            else if (tag == 282) dpiX = ReadTiffRational(data, valueOrOffset, littleEndian, dpiX);
            else if (tag == 283) dpiY = ReadTiffRational(data, valueOrOffset, littleEndian, dpiY);
            else if (tag == 296) unit = ReadTiffScalar(data, type, count, valueOrOffset, littleEndian);
        }

        if (unit == 3) {
            dpiX *= 2.54;
            dpiY *= 2.54;
        }

        info = new OfficeImageInfo(OfficeImageFormat.Tiff, width, height, dpiX, dpiY);
        return width > 0 && height > 0;
    }

    private static bool TryReadWebp(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 20 || GetAscii(data, 0, 4) != "RIFF" || GetAscii(data, 8, 4) != "WEBP") {
            return false;
        }

        int offset = 12;
        while (offset + 8 <= data.Length) {
            string chunkType = GetAscii(data, offset, 4);
            int chunkLength = ReadInt32LittleEndian(data, offset + 4);
            int payloadOffset = offset + 8;
            if (chunkLength < 0 || (long)payloadOffset + chunkLength > data.Length) return false;

            int width = 0;
            int height = 0;
            if (chunkType == "VP8L" && chunkLength >= 5 && data[payloadOffset] == 0x2F) {
                int packed = ReadInt32LittleEndian(data, payloadOffset + 1);
                width = (packed & 0x3FFF) + 1;
                height = ((packed >> 14) & 0x3FFF) + 1;
            } else if (chunkType == "VP8X" && chunkLength >= 10) {
                width = ReadUInt24LittleEndian(data, payloadOffset + 4) + 1;
                height = ReadUInt24LittleEndian(data, payloadOffset + 7) + 1;
            } else if (chunkType == "VP8 " && chunkLength >= 10 &&
                       data[payloadOffset + 3] == 0x9D && data[payloadOffset + 4] == 0x01 && data[payloadOffset + 5] == 0x2A) {
                width = ReadUInt16LittleEndian(data, payloadOffset + 6) & 0x3FFF;
                height = ReadUInt16LittleEndian(data, payloadOffset + 8) & 0x3FFF;
            }

            if (width > 0 && height > 0) {
                info = new OfficeImageInfo(OfficeImageFormat.Webp, width, height);
                return true;
            }

            long next = (long)payloadOffset + chunkLength + (chunkLength & 1);
            if (next > int.MaxValue) return false;
            offset = (int)next;
        }

        return false;
    }

    private static bool TryReadIcon(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 22) {
            return false;
        }

        int reserved = ReadUInt16LittleEndian(data, 0);
        int type = ReadUInt16LittleEndian(data, 2);
        int count = ReadUInt16LittleEndian(data, 4);
        if (reserved != 0 || type != 1 || count <= 0) {
            return false;
        }

        int width = data[6] == 0 ? 256 : data[6];
        int height = data[7] == 0 ? 256 : data[7];
        info = new OfficeImageInfo(OfficeImageFormat.Icon, width, height);
        return width > 0 && height > 0;
    }

    private static bool TryReadPcx(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 128 || data[0] != 0x0A || data[2] != 0x01) {
            return false;
        }

        int xMin = ReadUInt16LittleEndian(data, 4);
        int yMin = ReadUInt16LittleEndian(data, 6);
        int xMax = ReadUInt16LittleEndian(data, 8);
        int yMax = ReadUInt16LittleEndian(data, 10);
        int width = xMax - xMin + 1;
        int height = yMax - yMin + 1;
        double dpiX = ReadUInt16LittleEndian(data, 12);
        double dpiY = ReadUInt16LittleEndian(data, 14);

        info = new OfficeImageInfo(OfficeImageFormat.Pcx, width, height, dpiX, dpiY);
        return width > 0 && height > 0;
    }

    private static bool TryReadEmf(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 88) {
            return false;
        }

        int recordType = ReadInt32LittleEndian(data, 0);
        int headerSize = ReadInt32LittleEndian(data, 4);
        int signature = ReadInt32LittleEndian(data, 40);
        if (recordType != 1 || headerSize < 88 || headerSize > data.Length || signature != 0x464D4520) {
            return false;
        }

        int frameLeft = ReadInt32LittleEndian(data, 24);
        int frameTop = ReadInt32LittleEndian(data, 28);
        int frameRight = ReadInt32LittleEndian(data, 32);
        int frameBottom = ReadInt32LittleEndian(data, 36);
        int deviceWidth = ReadInt32LittleEndian(data, 72);
        int deviceHeight = ReadInt32LittleEndian(data, 76);
        int millimetersWidth = ReadInt32LittleEndian(data, 80);
        int millimetersHeight = ReadInt32LittleEndian(data, 84);

        double dpiX = millimetersWidth > 0 && deviceWidth > 0 ? deviceWidth * 25.4 / millimetersWidth : 96.0;
        double dpiY = millimetersHeight > 0 && deviceHeight > 0 ? deviceHeight * 25.4 / millimetersHeight : 96.0;
        int width = (int)Math.Round(Math.Abs(frameRight - frameLeft) / 2540.0 * dpiX);
        int height = (int)Math.Round(Math.Abs(frameBottom - frameTop) / 2540.0 * dpiY);

        if (width <= 0 || height <= 0) {
            int boundsLeft = ReadInt32LittleEndian(data, 8);
            int boundsTop = ReadInt32LittleEndian(data, 12);
            int boundsRight = ReadInt32LittleEndian(data, 16);
            int boundsBottom = ReadInt32LittleEndian(data, 20);
            width = Math.Abs(boundsRight - boundsLeft);
            height = Math.Abs(boundsBottom - boundsTop);
        }

        info = new OfficeImageInfo(OfficeImageFormat.Emf, width, height, dpiX, dpiY);
        return width > 0 && height > 0;
    }

    private static bool TryReadWmf(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 22 || ReadInt32LittleEndian(data, 0) != unchecked((int)0x9AC6CDD7)) {
            return false;
        }

        if (!HasValidPlaceableWmfChecksum(data)) {
            return false;
        }

        int left = ReadInt16LittleEndian(data, 6);
        int top = ReadInt16LittleEndian(data, 8);
        int right = ReadInt16LittleEndian(data, 10);
        int bottom = ReadInt16LittleEndian(data, 12);
        int unitsPerInch = ReadUInt16LittleEndian(data, 14);
        if (unitsPerInch <= 0) {
            return false;
        }

        int width = (int)Math.Round(Math.Abs(right - left) * 96.0 / unitsPerInch);
        int height = (int)Math.Round(Math.Abs(bottom - top) * 96.0 / unitsPerInch);
        info = new OfficeImageInfo(OfficeImageFormat.Wmf, width, height);
        return width > 0 && height > 0;
    }

    private static bool TryReadSvg(byte[] data, string? fileName, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        bool likelySvg = FromExtension(fileName) == OfficeImageFormat.Svg;
        if (!likelySvg) {
            var prefix = GetAscii(data, 0, Math.Min(data.Length, 256)).TrimStart();
            likelySvg = prefix.StartsWith("<svg", StringComparison.OrdinalIgnoreCase) ||
                        prefix.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase);
        }

        if (!likelySvg) {
            return false;
        }

        try {
            using var ms = new MemoryStream(data);
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            };
            using var reader = XmlReader.Create(ms, settings);
            var document = XDocument.Load(reader, LoadOptions.None);
            var root = document.Root;
            if (root == null || !root.Name.LocalName.Equals("svg", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            bool hasWidth = TryParseSvgLength(root.Attribute("width")?.Value, out double width);
            bool hasHeight = TryParseSvgLength(root.Attribute("height")?.Value, out double height);
            bool hasViewBox = TryParseSvgViewBox(root.Attribute("viewBox")?.Value, out double viewBoxWidth, out double viewBoxHeight);
            double? aspectRatio = hasViewBox ? viewBoxWidth / viewBoxHeight : (double?)null;
            if (!aspectRatio.HasValue && hasWidth && hasHeight) aspectRatio = width / height;

            if (hasWidth && !hasHeight && aspectRatio.HasValue) {
                height = width / aspectRatio.Value;
                hasHeight = true;
            } else if (!hasWidth && hasHeight && aspectRatio.HasValue) {
                width = height * aspectRatio.Value;
                hasWidth = true;
            } else if (!hasWidth && !hasHeight && hasViewBox) {
                width = viewBoxWidth;
                height = viewBoxHeight;
                hasWidth = true;
                hasHeight = true;
            }

            int pixelWidth = hasWidth ? Math.Max(1, (int)Math.Round(width)) : 0;
            int pixelHeight = hasHeight ? Math.Max(1, (int)Math.Round(height)) : 0;
            info = new OfficeImageInfo(OfficeImageFormat.Svg, pixelWidth, pixelHeight, 96D, 96D, aspectRatio);
            return true;
        } catch (XmlException) {
            info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
            return false;
        } catch (InvalidOperationException) {
            info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
            return false;
        } catch (IOException) {
            info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
            return false;
        } catch (ArgumentException) {
            info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
            return false;
        }
    }

    private static bool TryReadWebp(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 25 ||
            GetAscii(data, 0, 4) != "RIFF" ||
            GetAscii(data, 8, 4) != "WEBP") {
            return false;
        }

        string chunkType = GetAscii(data, 12, 4);
        int width;
        int height;
        if (chunkType == "VP8X" && data.Length >= 30) {
            width = 1 + ReadUInt24LittleEndian(data, 24);
            height = 1 + ReadUInt24LittleEndian(data, 27);
        } else if (chunkType == "VP8L" && data.Length >= 25 && data[20] == 0x2F) {
            width = 1 + data[21] + ((data[22] & 0x3F) << 8);
            height = 1 + ((data[22] & 0xC0) >> 6) + (data[23] << 2) + ((data[24] & 0x0F) << 10);
        } else if (chunkType == "VP8 " && data.Length >= 30 &&
            data[23] == 0x9D && data[24] == 0x01 && data[25] == 0x2A) {
            width = ReadUInt16LittleEndian(data, 26) & 0x3FFF;
            height = ReadUInt16LittleEndian(data, 28) & 0x3FFF;
        } else {
            return false;
        }

        info = new OfficeImageInfo(OfficeImageFormat.Webp, width, height);
        return width > 0 && height > 0;
    }

    private static bool TryParseSvgLength(string? value, out double result) {
        result = 0D;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string normalized = value!.Trim().ToLowerInvariant();
        int unitStart = normalized.Length;
        while (unitStart > 0 && (char.IsLetter(normalized[unitStart - 1]) || normalized[unitStart - 1] == '%')) unitStart--;
        string numberText = normalized.Substring(0, unitStart).Trim();
        string unit = normalized.Substring(unitStart);
        if (!double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
            || double.IsNaN(number)
            || double.IsInfinity(number)
            || number <= 0D) return false;

        double multiplier;
        switch (unit) {
            case "":
            case "px":
                multiplier = 1D;
                break;
            case "pt":
                multiplier = 96D / 72D;
                break;
            case "pc":
                multiplier = 16D;
                break;
            case "in":
                multiplier = 96D;
                break;
            case "cm":
                multiplier = 96D / 2.54D;
                break;
            case "mm":
                multiplier = 96D / 25.4D;
                break;
            case "q":
                multiplier = 96D / 101.6D;
                break;
            default:
                return false;
        }
        result = number * multiplier;
        return !double.IsNaN(result) && !double.IsInfinity(result) && result > 0D;
    }

    private static bool TryParseSvgViewBox(string? value, out double width, out double height) {
        width = 0D;
        height = 0D;
        if (string.IsNullOrWhiteSpace(value)) return false;
        string[] parts = value!.Split(new[] { ' ', '\t', '\r', '\n', '\f', ',' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 4
            || !double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double minX)
            || !double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double minY)
            || !double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out width)
            || !double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out height)
            || double.IsNaN(minX) || double.IsInfinity(minX)
            || double.IsNaN(minY) || double.IsInfinity(minY)
            || double.IsNaN(width) || double.IsInfinity(width)
            || double.IsNaN(height) || double.IsInfinity(height)
            || width <= 0D || height <= 0D) {
            width = 0D;
            height = 0D;
            return false;
        }
        return true;
    }

    private static bool StartsWith(byte[] data, byte[] prefix) {
        if (data.Length < prefix.Length) return false;
        for (int i = 0; i < prefix.Length; i++) {
            if (data[i] != prefix[i]) return false;
        }

        return true;
    }

    private static bool IsStartOfFrame(byte marker) =>
        marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;

    private static string GetAscii(byte[] data, int offset, int count) {
        if (offset < 0 || count <= 0 || offset >= data.Length) return string.Empty;
        count = Math.Min(count, data.Length - offset);
        return System.Text.Encoding.ASCII.GetString(data, offset, count);
    }

    private static int ReadTiffScalar(byte[] data, int type, int count, int valueOrOffset, bool littleEndian) {
        if (count <= 0) return 0;
        if (type == 3) {
            return littleEndian ? valueOrOffset & 0xFFFF : (valueOrOffset >> 16) & 0xFFFF;
        }

        if (type == 4) {
            return valueOrOffset;
        }

        return 0;
    }

    private static double ReadTiffRational(byte[] data, int offset, bool littleEndian, double fallback) {
        if (offset < 0 || offset + 8 > data.Length) return fallback;
        int numerator = ReadInt32(data, offset, littleEndian);
        int denominator = ReadInt32(data, offset + 4, littleEndian);
        return denominator != 0 ? (double)numerator / denominator : fallback;
    }

    private static int ReadInt32(byte[] data, int offset, bool littleEndian) =>
        littleEndian ? ReadInt32LittleEndian(data, offset) : ReadInt32BigEndian(data, offset);

    private static int ReadUInt16(byte[] data, int offset, bool littleEndian) =>
        littleEndian ? ReadUInt16LittleEndian(data, offset) : ReadUInt16BigEndian(data, offset);

    private static int ReadInt32LittleEndian(byte[] data, int offset) =>
        offset + 4 <= data.Length
            ? data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24)
            : 0;

    private static int ReadInt32BigEndian(byte[] data, int offset) =>
        offset + 4 <= data.Length
            ? (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3]
            : 0;

    private static int ReadUInt16LittleEndian(byte[] data, int offset) =>
        offset + 2 <= data.Length ? data[offset] | (data[offset + 1] << 8) : 0;

    private static int ReadUInt24LittleEndian(byte[] data, int offset) =>
        offset + 3 <= data.Length ? data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) : 0;

    private static int ReadUInt16BigEndian(byte[] data, int offset) =>
        offset + 2 <= data.Length ? (data[offset] << 8) | data[offset + 1] : 0;

    private static int ReadUInt24LittleEndian(byte[] data, int offset) =>
        offset + 3 <= data.Length ? data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) : 0;

    private static short ReadInt16LittleEndian(byte[] data, int offset) =>
        offset + 2 <= data.Length ? unchecked((short)(data[offset] | (data[offset + 1] << 8))) : (short)0;

    private static bool HasValidPlaceableWmfChecksum(byte[] data) {
        int checksum = 0;
        for (int offset = 0; offset < 20; offset += 2) {
            checksum ^= ReadUInt16LittleEndian(data, offset);
        }

        return checksum == ReadUInt16LittleEndian(data, 20);
    }
}
