using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    internal sealed class Container {
        internal int Width, Height, PixelFormat, Transform, ImageOffset, ImageLength, AlphaOffset, AlphaLength;
        internal double DpiX = 96D, DpiY = 96D;
        internal bool HasIcc;
        internal int IccOffset, IccLength;
        internal OfficeImageMetadataKinds MetadataKinds;
        internal bool HasColorRenderingMetadata;
        internal bool Premultiplied => PixelFormat == 0x10 || PixelFormat == 0x17 || PixelFormat == 0x1A;
        internal bool Gray => PixelFormat == 0x08 || PixelFormat == 0x0B || PixelFormat == 0x13 ||
            PixelFormat == 0x3E || PixelFormat == 0x3F || PixelFormat == 0x11;
        internal int BitDepth => PixelFormat switch {
            0x08 or 0x0C or 0x0D or 0x0F or 0x10 => 1,
            0x0B or 0x15 or 0x16 or 0x17 => 2,
            0x13 or 0x12 or 0x40 or 0x1D => 3,
            0x3E or 0x3B or 0x42 or 0x3A => 4,
            0x3F or 0x18 or 0x41 or 0x1E => 6,
            0x11 or 0x1B or 0x19 or 0x1A => 7,
            _ => 0
        };
        internal bool HasAlpha => PixelFormat == 0x0F || PixelFormat == 0x10 || PixelFormat == 0x16 || PixelFormat == 0x17 ||
            PixelFormat == 0x1D || PixelFormat == 0x3A || PixelFormat == 0x1E || PixelFormat == 0x19 || PixelFormat == 0x1A;
        internal FrameHeader Frame = new();
        internal FrameHeader? SeparateAlpha;
    }
    private readonly struct DirectoryField {
        internal DirectoryField(int type, uint count, int value) { Type = type; Count = count; Value = value; }
        internal int Type { get; }
        internal uint Count { get; }
        internal int Value { get; }
    }

    internal static Container ReadContainer(byte[] bytes, CancellationToken cancellation) {
        cancellation.ThrowIfCancellationRequested();
        if (bytes.Length < 8 || bytes.Length > 128 * 1024 * 1024 ||
            bytes[0] != 0x49 || bytes[1] != 0x49 || bytes[2] != 0xBC || bytes[3] != 1)
            throw new FormatException("JPEG-XR container signature or size is invalid.");
        uint ifd = Little32(bytes, 4);
        if (ifd < 8 || ifd > bytes.Length - 2 || (ifd & 1) != 0)
            throw new FormatException("JPEG-XR directory offset is invalid.");
        int count = Little16(bytes, (int)ifd);
        if (count == 0 || (long)ifd + 2L + count * 12L + 4L > bytes.Length)
            throw new FormatException("JPEG-XR directory is truncated.");
        var fields = new Dictionary<int, DirectoryField>();
        int previous = -1;
        for (int i = 0; i < count; i++) {
            if ((i & 255) == 0) cancellation.ThrowIfCancellationRequested();
            int entry = (int)ifd + 2 + i * 12;
            int tag = Little16(bytes, entry), type = Little16(bytes, entry + 2);
            uint elements = Little32(bytes, entry + 4);
            if (tag <= previous) throw new FormatException("JPEG-XR directory tags are not ordered uniquely.");
            previous = tag;
            int size = type switch { 1 or 2 or 6 or 7 => 1, 3 or 8 => 2, 4 or 9 or 11 => 4, 5 or 10 or 12 => 8, _ => 0 };
            if (size == 0) continue; // Unknown element types do not supply decoded fields.
            long byteCount = (long)elements * size;
            uint position = byteCount <= 4 ? (uint)(entry + 8) : Little32(bytes, entry + 8);
            if (position > bytes.Length || byteCount > bytes.Length - (long)position)
                throw new FormatException("JPEG-XR directory value range is invalid.");
            fields.Add(tag, new DirectoryField(type, elements, (int)position));
        }
        uint next = Little32(bytes, (int)ifd + 2 + count * 12);
        if (next != 0)
            throw new FormatException("JPEG-XR multiple image directories are outside the managed contract.");
        var container = new Container {
            Width = RequiredUnsigned(bytes, fields, 0xBC80), Height = RequiredUnsigned(bytes, fields, 0xBC81),
            ImageOffset = RequiredUnsigned(bytes, fields, 0xBCC0), ImageLength = RequiredUnsigned(bytes, fields, 0xBCC1)
        };
        if (fields.ContainsKey(0x8769)) container.MetadataKinds |= OfficeImageMetadataKinds.Exif;
        if (fields.ContainsKey(0x02BC)) container.MetadataKinds |= OfficeImageMetadataKinds.Xmp;
        if (fields.ContainsKey(0xBC82) || fields.ContainsKey(0xBC83)) container.MetadataKinds |= OfficeImageMetadataKinds.Resolution;
        if (fields.ContainsKey(0x010E) || fields.ContainsKey(0x013B) || fields.ContainsKey(0x8298)) container.MetadataKinds |= OfficeImageMetadataKinds.Comments;
        container.HasColorRenderingMetadata = fields.ContainsKey(0xBC05) || fields.ContainsKey(0xA001);
        if (container.Width < 1 || container.Height < 1 || (long)container.Width * container.Height > 50_000_000L)
            throw new FormatException("JPEG-XR container dimensions exceed the managed limit.");
        if (!fields.TryGetValue(0xBC01, out var pixel) || pixel.Type != 1 || pixel.Count != 16)
            throw new FormatException("JPEG-XR pixel format is missing or malformed.");
        byte[] formatPrefix = { 0x24, 0xC3, 0xDD, 0x6F, 0x03, 0x4E, 0xFE, 0x4B, 0xB1, 0x85, 0x3D, 0x77, 0x76, 0x8D, 0xC9 };
        for (int i = 0; i < formatPrefix.Length; i++)
            if (bytes[pixel.Value + i] != formatPrefix[i]) throw new FormatException("JPEG-XR pixel format is unsupported.");
        container.PixelFormat = bytes[pixel.Value + 15];
        if (container.BitDepth == 0) throw new FormatException("JPEG-XR pixel format is outside the RGB/gray contract.");
        if (fields.ContainsKey(0xBC02)) container.Transform = RequiredUnsigned(bytes, fields, 0xBC02);
        if (container.Transform > 7) container.Transform = 0;
        if (container.Transform != 0) container.MetadataKinds |= OfficeImageMetadataKinds.Orientation;
        if (fields.TryGetValue(0x8773, out var icc)) {
            container.HasIcc = true;
            container.MetadataKinds |= OfficeImageMetadataKinds.Icc;
            if ((icc.Type == 1 || icc.Type == 7) && icc.Count > 0 && icc.Count <= 4 * 1024 * 1024) {
                container.IccOffset = icc.Value; container.IccLength = (int)icc.Count;
            }
        }
        container.DpiX = ReadResolution(bytes, fields, 0xBC82); container.DpiY = ReadResolution(bytes, fields, 0xBC83);
        bool alphaOffset = fields.ContainsKey(0xBCC2), alphaCount = fields.ContainsKey(0xBCC3);
        if (alphaOffset != alphaCount || alphaOffset && !container.HasAlpha)
            throw new FormatException("JPEG-XR separate alpha declarations are inconsistent.");
        if (container.ImageLength == 0) container.ImageLength = bytes.Length - container.ImageOffset;
        ValidateCodestreamRange(bytes, container.ImageOffset, container.ImageLength);
        container.Frame = ReadHeader(bytes, container.ImageOffset, container.ImageLength, cancellation);
        if ((container.Gray ? 0 : 7) != container.Frame.OutputColor || container.BitDepth != container.Frame.BitDepth)
            throw new FormatException("JPEG-XR pixel format and codestream color/alpha semantics disagree.");
        if (container.Frame.Width != container.Width || container.Frame.Height != container.Height ||
            container.Frame.Alpha && !container.HasAlpha || !alphaOffset && container.HasAlpha != container.Frame.Alpha)
            throw new FormatException("JPEG-XR container and image-plane declarations are inconsistent.");
        if (alphaOffset) {
            container.AlphaOffset = RequiredUnsigned(bytes, fields, 0xBCC2);
            container.AlphaLength = RequiredUnsigned(bytes, fields, 0xBCC3);
            ValidateCodestreamRange(bytes, container.AlphaOffset, container.AlphaLength);
            if (container.Frame.Alpha || container.ImageOffset < container.AlphaOffset + container.AlphaLength &&
                container.AlphaOffset < container.ImageOffset + container.ImageLength)
                throw new FormatException("JPEG-XR primary and separate alpha ranges overlap.");
            container.SeparateAlpha = ReadHeader(bytes, container.AlphaOffset, container.AlphaLength, cancellation);
            if (container.SeparateAlpha.Width != container.Width || container.SeparateAlpha.Height != container.Height ||
                container.SeparateAlpha.OutputColor != 0 || container.SeparateAlpha.Alpha ||
                container.SeparateAlpha.BitDepth != container.BitDepth)
                throw new FormatException("JPEG-XR separate alpha frame is inconsistent.");
        }
        // Association is defined by PIXEL_FORMAT. Legacy streams may leave the
        // flag clear; a primary stream without alpha has no meaningful alpha flag.
        FrameHeader? alphaHeader = container.SeparateAlpha ?? (container.Frame.Alpha ? container.Frame : null);
        if (alphaHeader?.Premultiplied == true && !container.Premultiplied)
            throw new FormatException("JPEG-XR alpha association contradicts the pixel format.");
        return container;
    }

    private static void ValidateCodestreamRange(byte[] bytes, int offset, int length) {
        if (offset < 8 || length < 16 || offset > bytes.Length - length)
            throw new FormatException("JPEG-XR declared codestream range is invalid.");
    }
    private static int RequiredUnsigned(byte[] bytes, Dictionary<int, DirectoryField> fields, int tag) {
        if (!fields.TryGetValue(tag, out var field) || field.Count != 1 || (field.Type != 1 && field.Type != 3 && field.Type != 4))
            throw new FormatException("JPEG-XR scalar field is missing or malformed.");
        uint value = field.Type == 1 ? bytes[field.Value] : field.Type == 3 ? (uint)Little16(bytes, field.Value) : Little32(bytes, field.Value);
        if (value > int.MaxValue) throw new FormatException("JPEG-XR scalar exceeds the managed range.");
        return (int)value;
    }
    private static double ReadResolution(byte[] bytes, Dictionary<int, DirectoryField> fields, int tag) {
        if (!fields.TryGetValue(tag, out var field)) return 96D;
        if (field.Type != 11 || field.Count != 1) throw new FormatException("JPEG-XR resolution field is malformed.");
        double value = BitConverter.ToSingle(BitConverter.GetBytes(Little32(bytes, field.Value)), 0);
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0)
            throw new FormatException("JPEG-XR resolution value is invalid.");
        return value == 0 ? 96D : value;
    }
    private static int Little16(byte[] bytes, int offset) => bytes[offset] | bytes[offset + 1] << 8;
    private static uint Little32(byte[] bytes, int offset) => (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);
}
