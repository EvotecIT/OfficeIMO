using System;
using System.Buffers.Binary;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using OfficeIMO.Drawing.Wic.Interop;

namespace OfficeIMO.Drawing.Wic;

/// <summary>Decodes HEIC using Microsoft's installed Windows Imaging Component HEIF decoder.</summary>
/// <remarks>No decoder binaries are bundled. Each call owns its COM objects and may run concurrently.
/// Core supplies single-item HEIC bytes without container transforms or color properties.</remarks>
public sealed partial class OfficeWicCodec : IOfficeRasterImageCodec, IOfficeRasterImageCodecAvailability {
    private const int MaximumInputBytes = 128 * 1024 * 1024;
    private const long MaximumPixels = 50_000_000;
    private const long MaximumWorkingBytes = 256L * 1024 * 1024;
    private const long WorkingOverheadBytes = 64L * 1024;

    /// <summary>Whether this operating system supports WIC. Installed extensions are checked separately.</summary>
    [SupportedOSPlatformGuard("windows")]
    public static bool IsSupported => OperatingSystem.IsWindows();

    /// <summary>Installs this codec as the process-wide platform fallback on Windows.</summary>
    /// <remarks>Does nothing on other operating systems. Does not install or download Windows extensions.</remarks>
    public static void Register() {
        if (IsSupported) OfficeRasterImageCodecs.Platform = new OfficeWicCodec();
    }

    /// <summary>Attempts to decode the first HEIC frame into tightly packed straight-alpha RGBA bytes.</summary>
    /// <remarks>Returns false for unsupported, corrupt, unavailable, or over-budget input.
    /// The byte signature, rather than the optional content type, selects HEIF. Output is 8 bits per channel.
    /// No color-management transform, gamma correction, orientation, or tone mapping is requested.</remarks>
    public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
        image = null;
        if (!IsSupported || encodedBytes == null || encodedBytes.Length > MaximumInputBytes || !IsHeif(encodedBytes)) return false;
        try {
            image = DecodeWindows(encodedBytes);
            return image != null;
        } catch (COMException) {
            return false;
        }
    }

    [SupportedOSPlatform("windows")]
    private static unsafe OfficeRasterImage? DecodeWindows(byte[] bytes) {
        using var apartment = new WicApartment();
        using var factory = WicNative.CreateFactory();
        // WIC holds the input pointer, including during lazy CopyPixels. Release every dependent
        // native reference inside the fixed block before allowing this memory to move.
        fixed (byte* input = bytes) {
            WicNative.Check(factory.Value.CreateStream(out nint streamPointer));
            using var stream = new WicObject<IWicStream>(streamPointer);
            WicNative.Check(stream.Value.InitializeFromMemory((nint)input, (uint)bytes.Length));
            // Activate only Microsoft's HEIF decoder; arbitrary third-party WIC decoders are not fallback engines.
            WicNative.Check(WicNative.CoCreateInstance(WicNative.HeifDecoderClass, 0, 1,
                WicNative.DecoderInterface, out nint decoderPointer));
            using var decoder = new WicObject<IWicBitmapDecoder>(decoderPointer);
            WicNative.Check(decoder.Value.Initialize(stream.Pointer, 0)); // WICDecodeMetadataCacheOnDemand
            WicNative.Check(decoder.Value.GetFrameCount(out uint count));
            if (count == 0) return null;
            WicNative.Check(decoder.Value.GetFrame(0, out nint framePointer));
            using var frame = new WicObject<IWicBitmapSource>(framePointer);
            WicNative.Check(frame.Value.GetSize(out uint width, out uint height));
            if (!FitsBudget(width, height, bytes.Length)) return null;

            WicNative.Check(factory.Value.CreateFormatConverter(out nint converterPointer));
            using var converter = new WicObject<IWicFormatConverter>(converterPointer);
            WicNative.Check(frame.Value.GetPixelFormat(out Guid sourceFormat));
            WicNative.Check(converter.Value.CanConvert(sourceFormat, WicNative.Rgba32, out int canConvert));
            if (canConvert == 0) return null;
            // A pixel format converter is not an IWICColorTransform. Keep source RGB/gamma;
            // the installed HEVC decoder still owns the necessary YUV-to-RGB reconstruction.
            WicNative.Check(converter.Value.Initialize(frame.Pointer, WicNative.Rgba32, 0, 0, 0, 0));
            int stride = checked((int)width * 4);
            byte[] rgba = new byte[checked(stride * (int)height)];
            fixed (byte* pixels = rgba) {
                WicNative.Check(converter.Value.CopyPixels(0, (uint)stride, (uint)rgba.Length, (nint)pixels));
            }
            return OfficeRasterImage.FromRgba32((int)width, (int)height, rgba);
        }
    }

    private static bool FitsBudget(uint width, uint height, int encodedBytes) {
        if (width == 0 || height == 0 || width > int.MaxValue / 4 || height > int.MaxValue) return false;
        ulong pixels = (ulong)width * height;
        if (pixels > MaximumPixels) return false;
        // Two managed RGBA buffers (FromRgba32 copies) plus an 8-byte/pixel allowance
        // for native source/conversion storage. WIC offers no hard native allocator quota.
        return (long)pixels * 16L + encodedBytes + WorkingOverheadBytes <= MaximumWorkingBytes;
    }

    private static bool IsHeif(ReadOnlySpan<byte> bytes) {
        if (bytes.Length < 16 || !bytes.Slice(4, 4).SequenceEqual("ftyp"u8)) return false;
        ulong size = BinaryPrimitives.ReadUInt32BigEndian(bytes);
        int header = 8;
        if (size == 1) {
            if (bytes.Length < 24) return false;
            size = BinaryPrimitives.ReadUInt64BigEndian(bytes.Slice(8));
            header = 16;
        } else if (size == 0) {
            size = (ulong)bytes.Length;
        }
        if (size < (ulong)(header + 8) || size > (ulong)bytes.Length || (size - (uint)header) % 4 != 0) return false;
        for (int i = header; i < (int)size; i += 4) {
            if (i == header + 4) continue; // Minor version is not a brand.
            ReadOnlySpan<byte> brand = bytes.Slice(i, 4);
            if (brand.SequenceEqual("heic"u8) || brand.SequenceEqual("heix"u8) ||
                brand.SequenceEqual("hevc"u8) || brand.SequenceEqual("hevx"u8) ||
                brand.SequenceEqual("heim"u8) || brand.SequenceEqual("heis"u8) ||
                brand.SequenceEqual("hevm"u8) || brand.SequenceEqual("hevs"u8) || brand.SequenceEqual("mif1"u8)) return true;
        }
        return false;
    }
}
