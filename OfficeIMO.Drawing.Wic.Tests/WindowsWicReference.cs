using System;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using OfficeIMO.Drawing.Wic.Interop;

namespace OfficeIMO.Drawing.Wic.Tests;

/// <summary>Test-only full-file WIC pipeline with native metadata and sRGB color management.</summary>
/// <remarks>Shares ABI declarations with the codec, but never calls its decoder or Core's HEIF parser,
/// compositor, orientation or color conversion. WIC receives the original complete file.</remarks>
[SupportedOSPlatform("windows")]
internal sealed partial class WindowsWicReference {
    private readonly byte[] _pixels;
    private readonly int _storedWidth;
    private readonly int _storedHeight;
    private readonly ushort _orientation;
    internal int Width => _orientation >= 5 ? _storedHeight : _storedWidth;
    internal int Height => _orientation >= 5 ? _storedWidth : _storedHeight;

    private WindowsWicReference(byte[] pixels, uint width, uint height, ushort orientation) {
        _pixels = pixels;
        _storedWidth = (int)width;
        _storedHeight = (int)height;
        _orientation = orientation;
    }

    internal OfficeColor GetPixel(int x, int y) {
        // Use only WIC's reported orientation, independently of Core's HEIF transform interpretation.
        (int sx, int sy) = _orientation switch {
            2 => (_storedWidth - 1 - x, y),
            3 => (_storedWidth - 1 - x, _storedHeight - 1 - y),
            4 => (x, _storedHeight - 1 - y),
            5 => (y, x),
            6 => (y, _storedHeight - 1 - x),
            7 => (_storedWidth - 1 - y, _storedHeight - 1 - x),
            8 => (_storedWidth - 1 - y, x),
            _ => (x, y)
        };
        int offset = (sy * _storedWidth + sx) * 4;
        return OfficeColor.FromRgba(_pixels[offset], _pixels[offset + 1], _pixels[offset + 2], _pixels[offset + 3]);
    }

    internal static unsafe bool TryDecode(byte[] bytes, out WindowsWicReference? reference, out string detail) {
        reference = null;
        detail = string.Empty;
        if (bytes.Length == 0 || bytes.Length > 32 * 1024 * 1024) {
            detail = "Full-file reference requires at most 32 MiB encoded input.";
            return false;
        }
        try {
            using var apartment = new WicApartment();
            using var factory = WicNative.CreateFactory();
            fixed (byte* input = bytes) {
                WicNative.Check(factory.Value.CreateStream(out nint streamPointer));
                using var stream = new WicObject<IWicStream>(streamPointer);
                WicNative.Check(stream.Value.InitializeFromMemory((nint)input, (uint)bytes.Length));
                WicNative.Check(WicNative.CoCreateInstance(WicNative.HeifDecoderClass, 0, 1,
                    WicNative.DecoderInterface, out nint decoderPointer));
                using var decoder = new WicObject<IWicBitmapDecoder>(decoderPointer);
                WicNative.Check(decoder.Value.Initialize(stream.Pointer, 0));
                WicNative.Check(decoder.Value.GetFrame(0, out nint framePointer));
                using var frame = new WicObject<IWicReferenceFrame>(framePointer);
                WicNative.Check(frame.Value.GetSize(out uint width, out uint height));
                ulong pixels = (ulong)width * height;
                // Retained Core raster + reference pixels + native allowance, without pixel snapshots.
                if (width == 0 || height == 0 || pixels > 16_000_000 ||
                    pixels * 16 + (ulong)bytes.Length + 65536 > 256UL * 1024 * 1024) {
                    detail = "Full-file reference exceeds its 16 MP / 256 MiB working-memory admission budget.";
                    return false;
                }
                ushort orientation = ReadOrientation(frame.Value);
                WicNative.Check(frame.Value.GetColorContexts(0, 0, out uint contexts));
                WicNative.Check(factory.Value.CreateFormatConverter(out nint converterPointer));
                using var converter = new WicObject<IWicFormatConverter>(converterPointer);
                WicNative.Check(converter.Value.Initialize(frame.Pointer, WicNative.Rgba32, 0, 0, 0, 0));

                byte[] rgba = new byte[checked((int)pixels * 4)];
                if (contexts == 0) {
                    // Microsoft's WIC contract infers sRGB for unsigned RGB without a color context.
                    CopyPixels(converter.Value, width, rgba);
                } else {
                    WicNative.Check(factory.Value.CreateColorContext(out nint sourcePointer));
                    using var source = new WicObject<IWicReferenceColorContext>(sourcePointer);
                    nint context = source.Pointer;
                    WicNative.Check(frame.Value.GetColorContexts(1, (nint)(&context), out _));
                    WicNative.Check(factory.Value.CreateColorContext(out nint targetPointer));
                    using var target = new WicObject<IWicReferenceColorContext>(targetPointer);
                    WicNative.Check(target.Value.InitializeFromExifColorSpace(1)); // EXIF sRGB
                    WicNative.Check(factory.Value.CreateColorTransformer(out nint transformPointer));
                    using var transform = new WicObject<IWicReferenceColorTransform>(transformPointer);
                    WicNative.Check(transform.Value.Initialize(converter.Pointer, source.Pointer, target.Pointer, WicNative.Rgba32));
                    CopyPixels(transform.Value, width, rgba);
                }
                reference = new WindowsWicReference(rgba, width, height, orientation);
                detail = $"Native full-file WIC: stored {width}x{height}, orientation {orientation}, color contexts {contexts}; sRGB output.";
                return true;
            }
        } catch (COMException ex) {
            detail = $"Full-file WIC reference unavailable: HRESULT 0x{ex.HResult:X8}.";
            return false;
        }
    }

    private static ushort ReadOrientation(IWicReferenceFrame frame) {
        WicNative.Check(frame.GetMetadataQueryReader(out nint pointer));
        using var reader = new WicObject<IWicReferenceMetadata>(pointer);
        // WICHeifOrientation uses System.Photo.Orientation numbering, independently of EXIF.
        WicPropVariant value = default;
        int hr = reader.Value.GetMetadataByName("/heif/{ushort=1}", ref value);
        try {
            if (hr == unchecked((int)0x88982F40)) return 1; // WINCODEC_ERR_PROPERTYNOTFOUND
            WicNative.Check(hr);
            if (value.Type != 18 || value.UInt16 is < 1 or > 8) // VT_UI2
                throw new COMException("WIC did not provide a usable HEIF orientation.", unchecked((int)0x80004005));
            return value.UInt16;
        } finally {
            PropVariantClear(ref value);
        }
    }

    private static unsafe void CopyPixels(IWicBitmapSource source, uint width, byte[] rgba) {
        fixed (byte* pixels = rgba) {
            WicNative.Check(source.CopyPixels(0, checked(width * 4), (uint)rgba.Length, (nint)pixels));
        }
    }

    [LibraryImport("ole32.dll")]
    private static partial int PropVariantClear(ref WicPropVariant value);
}
