using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;
using OfficeIMO.Drawing;

namespace OfficeIMO.Drawing.ImageIO.Tests;

/// <summary>Test-only system encoder and full-file color-managed reference. This file can be linked into verification programs.</summary>
/// <remarks>Independent of the production decoder: the reference keeps the source color space and asks ImageIO
/// for a transformed full-size primary image, then draws into sRGB. No Core composition/color code is used.</remarks>
public static class AppleImageIOReference {
    /// <summary>Describes the native stored image before any bitmap drawing or destination color matching.</summary>
    public static string DescribeStoredImage(byte[] bytes) {
        RequireMac();
        if (bytes == null || bytes.Length == 0 || bytes.Length > 32 * 1024 * 1024) throw new ArgumentException("Invalid probe input.");
        using var symbols = new Symbols();
        IntPtr no = symbols.Foundation("kCFBooleanFalse");
        using var options = new Owned(N.CFDictionaryCreate(IntPtr.Zero,
            new[] { symbols.ImageIO("kCGImageSourceShouldCache"), symbols.ImageIO("kCGImageSourceShouldAllowFloat") },
            new[] { no, no }, 2, IntPtr.Zero, IntPtr.Zero));
        using var data = new Owned(N.CFDataCreate(IntPtr.Zero, bytes, bytes.Length));
        using var source = new Owned(N.CGImageSourceCreateWithData(data.Value, options.Value));
        if (source.IsInvalid) throw new ArgumentException("Invalid image source.");
        using var properties = new Owned(N.CGImageSourceCopyPropertiesAtIndex(source.Value, 0, options.Value));
        long width = ReadDimension(properties.Value, symbols.ImageIO("kCGImagePropertyPixelWidth"));
        long height = ReadDimension(properties.Value, symbols.ImageIO("kCGImagePropertyPixelHeight"));
        if (width > 16_000_000 || height > 16_000_000 || width * height > 16_000_000) throw new ArgumentException("Probe pixel budget exceeded.");
        using var image = new Owned(N.CGImageSourceCreateImageAtIndex(source.Value, 0, options.Value));
        if (image.IsInvalid) throw new ArgumentException("ImageIO rejected the image.");
        IntPtr space = N.CGImageGetColorSpace(image.Value);
        using var name = new Owned(space == IntPtr.Zero ? IntPtr.Zero : N.CGColorSpaceCopyName(space));
        var text = new StringBuilder(512);
        if (!name.IsInvalid) N.CFStringGetCString(name.Value, text, text.Capacity, 0x08000100);
        string spaceName = text.Length == 0 ? "<unnamed>" : text.ToString();
        IntPtr profileName = N.CFDictionaryGetValue(properties.Value, symbols.ImageIO("kCGImagePropertyProfileName"));
        text.Clear();
        if (profileName != IntPtr.Zero && N.CFGetTypeID(profileName) == N.CFStringGetTypeID())
            N.CFStringGetCString(profileName, text, text.Capacity, 0x08000100);
        return $"{N.CGImageGetWidth(image.Value)}x{N.CGImageGetHeight(image.Value)}, colorSpace={spaceName}, profile={text}, " +
            $"model={(space == IntPtr.Zero ? -1 : N.CGColorSpaceGetModel(space))}, " +
            $"bpc={N.CGImageGetBitsPerComponent(image.Value)}, bpp={N.CGImageGetBitsPerPixel(image.Value)}, " +
            $"bitmapInfo=0x{N.CGImageGetBitmapInfo(image.Value):X}, alphaInfo={N.CGImageGetAlphaInfo(image.Value)}";
    }
    /// <summary>Creates an original opaque 32 by 24 pixel color pattern, without camera metadata.</summary>
    public static byte[] ColorfulRgba() {
        byte[] rgba = new byte[32 * 24 * 4];
        for (int y = 0; y < 24; y++) {
            for (int x = 0; x < 32; x++) {
                int i = (y * 32 + x) * 4;
                rgba[i] = (byte)(x < 16 ? 224 : 32);
                rgba[i + 1] = (byte)(y < 12 ? 64 : 208);
                rgba[i + 2] = (byte)((x / 8 + y / 6) % 2 == 0 ? 192 : 48);
                rgba[i + 3] = 255;
            }
        }
        return rgba;
    }

    /// <summary>Encodes the original 32 by 24 pattern with Apple's HEIC encoder. No file is written.</summary>
    public static byte[] CreateColorfulHeic(bool displayP3 = false, int orientation = 1) =>
        Encode(ColorfulRgba(), 32, 24, "public.heic", displayP3, orientation);

    /// <summary>Writes a fixture explicitly requested by the caller; refuses to overwrite an existing file.</summary>
    public static void WriteColorfulHeic(string path) {
        byte[] bytes = CreateColorfulHeic();
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write);
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>Encodes supplied straight RGBA samples using a system encoder and an explicit RGB space.</summary>
    /// <remarks>Use public.png for exact sample controls, public.heic for the lossy HEVC codec probe.
    /// Output bytes can vary with OS/encoder version. Original generated patterns can be retained as MIT fixtures.</remarks>
    public static byte[] Encode(byte[] rgba, int width, int height, string type = "public.heic", bool displayP3 = false, int orientation = 1) =>
        EncodeSamples(rgba, width, height, type, displayP3, orientation, false);

    /// <summary>Encodes interleaved straight gray/alpha pairs for monochrome native sample tests.</summary>
    public static byte[] EncodeGrayPng(byte[] grayAlpha, int width, int height) =>
        EncodeSamples(grayAlpha, width, height, "public.png", false, 1, true);

    private static byte[] EncodeSamples(byte[] rgba, int width, int height, string type, bool displayP3, int orientation, bool gray) {
        RequireMac();
        int components = gray ? 2 : 4;
        if (width <= 0 || height <= 0 || (long)width * height > 1_000_000 || rgba == null ||
            rgba.LongLength != (long)width * height * components || orientation < 1 || orientation > 8)
            throw new ArgumentException("Invalid fixture dimensions, samples or orientation.");
        using var symbols = new Symbols();
        using var space = new Owned(gray ? N.CGColorSpaceCreateDeviceGray() :
            N.CGColorSpaceCreateWithName(symbols.Graphics(displayP3 ? "kCGColorSpaceDisplayP3" : "kCGColorSpaceSRGB")));
        using var data = new Owned(N.CFDataCreate(IntPtr.Zero, rgba, rgba.Length));
        using var provider = new Owned(N.CGDataProviderCreateWithCFData(data.Value));
        using var image = new Owned(N.CGImageCreate((nuint)width, (nuint)height, 8, (nuint)components * 8, (nuint)width * (nuint)components,
            space.Value, gray ? 3U : 0x4003, provider.Value, IntPtr.Zero, false, 0)); // alphaLast (straight)
        using var target = new Owned(N.CFDataCreateMutable(IntPtr.Zero, 0));
        using var uti = new Owned(N.CFStringCreateWithCString(IntPtr.Zero, type, 0x08000100));
        using var destination = new Owned(N.CGImageDestinationCreateWithData(target.Value, uti.Value, 1, IntPtr.Zero));
        if (image.IsInvalid || target.IsInvalid || destination.IsInvalid)
            throw new InvalidOperationException("System image encoder unavailable.");
        double quality = 1;
        long orientationValue = orientation;
        using var qualityNumber = new Owned(N.CFNumberCreateDouble(IntPtr.Zero, 6, ref quality)); // Float64
        using var orientationNumber = new Owned(N.CFNumberCreateInt64(IntPtr.Zero, 4, ref orientationValue));
        using var properties = new Owned(N.CFDictionaryCreate(IntPtr.Zero,
            new[] { symbols.ImageIO("kCGImageDestinationLossyCompressionQuality"), symbols.ImageIO("kCGImagePropertyOrientation") },
            new[] { qualityNumber.Value, orientationNumber.Value }, 2, IntPtr.Zero, IntPtr.Zero));
        N.CGImageDestinationAddImage(destination.Value, image.Value, properties.Value);
        if (!N.CGImageDestinationFinalize(destination.Value)) throw new InvalidOperationException("System image encoder failed.");
        long length = N.CFDataGetLength(target.Value);
        if (length < 1 || length > 16 * 1024 * 1024) throw new InvalidOperationException("Unexpected fixture size.");
        byte[] result = new byte[(int)length];
        Marshal.Copy(N.CFDataGetBytePtr(target.Value), result, 0, result.Length);
        return result;
    }

    /// <summary>Decodes the full primary image with native container/orientation transforms and sRGB matching.</summary>
    /// <remarks>Uses the original image, never an embedded preview. The output budget is checked before decoding.
    /// This deliberately differs from the raw tile codec and serves as an independent composition reference.</remarks>
    public static OfficeRasterImage DecodeFullFile(byte[] bytes, long maximumPixels = 4_000_000) {
        RequireMac();
        if (bytes == null || bytes.Length == 0 || bytes.Length > 32 * 1024 * 1024 || maximumPixels < 1 || maximumPixels > 16_000_000)
            throw new ArgumentException("Reference input exceeds its test budget.");
        using var symbols = new Symbols();
        IntPtr no = symbols.Foundation("kCFBooleanFalse"), yes = symbols.Foundation("kCFBooleanTrue");
        using var options = new Owned(N.CFDictionaryCreate(IntPtr.Zero,
            new[] { symbols.ImageIO("kCGImageSourceShouldCache"), symbols.ImageIO("kCGImageSourceShouldAllowFloat") },
            new[] { no, no }, 2, IntPtr.Zero, IntPtr.Zero));
        using var data = new Owned(N.CFDataCreate(IntPtr.Zero, bytes, bytes.Length));
        using var source = new Owned(N.CGImageSourceCreateWithData(data.Value, options.Value));
        if (source.IsInvalid) throw new InvalidOperationException("ImageIO rejected the reference input.");
        nuint index = N.CGImageSourceGetPrimaryImageIndex(source.Value);
        using var properties = new Owned(N.CGImageSourceCopyPropertiesAtIndex(source.Value, index, options.Value));
        long width = ReadDimension(properties.Value, symbols.ImageIO("kCGImagePropertyPixelWidth"));
        long height = ReadDimension(properties.Value, symbols.ImageIO("kCGImagePropertyPixelHeight"));
        if (width > maximumPixels || height > maximumPixels || width * height > maximumPixels)
            throw new ArgumentException("Reference image exceeds its pixel budget.");
        long maxDimension = Math.Max(width, height);
        using var size = new Owned(N.CFNumberCreateInt64(IntPtr.Zero, 4, ref maxDimension));
        using var thumbnailOptions = new Owned(N.CFDictionaryCreate(IntPtr.Zero,
            new[] { symbols.ImageIO("kCGImageSourceCreateThumbnailFromImageAlways"), symbols.ImageIO("kCGImageSourceCreateThumbnailWithTransform"),
                symbols.ImageIO("kCGImageSourceThumbnailMaxPixelSize"), symbols.ImageIO("kCGImageSourceShouldAllowFloat") },
            new[] { yes, yes, size.Value, no }, 4, IntPtr.Zero, IntPtr.Zero));
        using var image = new Owned(N.CGImageSourceCreateThumbnailAtIndex(source.Value, index, thumbnailOptions.Value));
        if (image.IsInvalid) throw new InvalidOperationException("Native full-file reference decode failed.");
        int actualWidth = checked((int)N.CGImageGetWidth(image.Value));
        int actualHeight = checked((int)N.CGImageGetHeight(image.Value));
        if (actualWidth < 1 || actualHeight < 1 || (long)actualWidth * actualHeight > maximumPixels)
            throw new ArgumentException("Unexpected reference dimensions.");
        using var srgb = new Owned(N.CGColorSpaceCreateWithName(symbols.Graphics("kCGColorSpaceSRGB")));
        byte[] rgba = new byte[checked(actualWidth * actualHeight * 4)];
        GCHandle pin = GCHandle.Alloc(rgba, GCHandleType.Pinned);
        try {
            using var context = new Owned(N.CGBitmapContextCreate(pin.AddrOfPinnedObject(), (nuint)actualWidth,
                (nuint)actualHeight, 8, (nuint)actualWidth * 4, srgb.Value, 0x4001));
            if (context.IsInvalid) throw new InvalidOperationException("Native reference context unavailable.");
            N.CGContextSetBlendMode(context.Value, 17);
            N.CGContextSetInterpolationQuality(context.Value, 1);
            N.CGContextDrawImage(context.Value, new N.Rect(actualWidth, actualHeight), image.Value);
            N.CGContextFlush(context.Value);
        } finally { pin.Free(); }
        for (int i = 0; i < rgba.Length; i += 4) {
            int a = rgba[i + 3];
            if (a == 255) continue;
            for (int c = 0; c < 3; c++) rgba[i + c] = a == 0 ? (byte)0 : (byte)Math.Min(255, (rgba[i + c] * 255 + a / 2) / a);
        }
        return OfficeRasterImage.FromRgba32(actualWidth, actualHeight, rgba);
    }

    private static long ReadDimension(IntPtr properties, IntPtr key) {
        if (properties == IntPtr.Zero) throw new ArgumentException("Missing image properties.");
        IntPtr number = N.CFDictionaryGetValue(properties, key);
        if (number == IntPtr.Zero || N.CFGetTypeID(number) != N.CFNumberGetTypeID() ||
            !N.CFNumberGetValue(number, 4, out long value) || value < 1 || value > int.MaxValue)
            throw new ArgumentException("Invalid reference dimensions.");
        return value;
    }

    private static void RequireMac() {
        if (!OperatingSystem.IsMacOSVersionAtLeast(10, 13)) throw new PlatformNotSupportedException("The test reference requires macOS 10.13+.");
    }

    private sealed class Owned : SafeHandleZeroOrMinusOneIsInvalid {
        internal Owned(IntPtr value) : base(true) => SetHandle(value);
        internal IntPtr Value => handle;
        protected override bool ReleaseHandle() { N.CFRelease(handle); return true; }
    }

    private sealed class Symbols : IDisposable {
        private IntPtr _cf, _cg, _io;
        internal Symbols() {
            try { _cf = NativeLibrary.Load(N.CF); _cg = NativeLibrary.Load(N.CG); _io = NativeLibrary.Load(N.IO); }
            catch { Dispose(); throw; }
        }
        internal IntPtr Foundation(string name) => Marshal.ReadIntPtr(NativeLibrary.GetExport(_cf, name));
        internal IntPtr Graphics(string name) => Marshal.ReadIntPtr(NativeLibrary.GetExport(_cg, name));
        internal IntPtr ImageIO(string name) => Marshal.ReadIntPtr(NativeLibrary.GetExport(_io, name));
        public void Dispose() {
            if (_io != IntPtr.Zero) { NativeLibrary.Free(_io); _io = IntPtr.Zero; }
            if (_cg != IntPtr.Zero) { NativeLibrary.Free(_cg); _cg = IntPtr.Zero; }
            if (_cf != IntPtr.Zero) { NativeLibrary.Free(_cf); _cf = IntPtr.Zero; }
        }
    }

    private static class N {
        internal const string CF = "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation";
        internal const string CG = "/System/Library/Frameworks/CoreGraphics.framework/CoreGraphics";
        internal const string IO = "/System/Library/Frameworks/ImageIO.framework/ImageIO";
        [StructLayout(LayoutKind.Sequential)]
        internal readonly struct Rect {
            internal readonly double X, Y, Width, Height;
            internal Rect(int width, int height) { X = 0; Y = 0; Width = width; Height = height; }
        }
        [DllImport(CF)] internal static extern void CFRelease(IntPtr value);
        [DllImport(CF)] internal static extern IntPtr CFDataCreate(IntPtr allocator, byte[] bytes, nint length);
        [DllImport(CF)] internal static extern IntPtr CFDataCreateMutable(IntPtr allocator, nint capacity);
        [DllImport(CF)] internal static extern nint CFDataGetLength(IntPtr data);
        [DllImport(CF)] internal static extern IntPtr CFDataGetBytePtr(IntPtr data);
        [DllImport(CF)] internal static extern IntPtr CFDictionaryCreate(IntPtr allocator, IntPtr[] keys, IntPtr[] values, nint count, IntPtr keyCallbacks, IntPtr valueCallbacks);
        [DllImport(CF)] internal static extern IntPtr CFDictionaryGetValue(IntPtr dictionary, IntPtr key);
        [DllImport(CF)] internal static extern nuint CFGetTypeID(IntPtr value);
        [DllImport(CF)] internal static extern nuint CFNumberGetTypeID();
        [DllImport(CF)] internal static extern nuint CFStringGetTypeID();
        [DllImport(CF)] [return: MarshalAs(UnmanagedType.I1)] internal static extern bool CFNumberGetValue(IntPtr number, int type, out long value);
        [DllImport(CF, EntryPoint = "CFNumberCreate")] internal static extern IntPtr CFNumberCreateDouble(IntPtr allocator, int type, ref double value);
        [DllImport(CF, EntryPoint = "CFNumberCreate")] internal static extern IntPtr CFNumberCreateInt64(IntPtr allocator, int type, ref long value);
        [DllImport(CF)] internal static extern IntPtr CFStringCreateWithCString(IntPtr allocator, [MarshalAs(UnmanagedType.LPUTF8Str)] string value, uint encoding);
        [DllImport(CF)] [return: MarshalAs(UnmanagedType.I1)] internal static extern bool CFStringGetCString(IntPtr value, StringBuilder buffer, nint size, uint encoding);
        [DllImport(CG)] internal static extern IntPtr CGColorSpaceCreateWithName(IntPtr name);
        [DllImport(CG)] internal static extern IntPtr CGColorSpaceCreateDeviceGray();
        [DllImport(CG)] internal static extern IntPtr CGColorSpaceCopyName(IntPtr space);
        [DllImport(CG)] internal static extern int CGColorSpaceGetModel(IntPtr space);
        [DllImport(CG)] internal static extern IntPtr CGImageGetColorSpace(IntPtr image);
        [DllImport(CG)] internal static extern nuint CGImageGetBitsPerComponent(IntPtr image);
        [DllImport(CG)] internal static extern nuint CGImageGetBitsPerPixel(IntPtr image);
        [DllImport(CG)] internal static extern uint CGImageGetBitmapInfo(IntPtr image);
        [DllImport(CG)] internal static extern uint CGImageGetAlphaInfo(IntPtr image);
        [DllImport(CG)] internal static extern IntPtr CGDataProviderCreateWithCFData(IntPtr data);
        [DllImport(CG)] internal static extern IntPtr CGImageCreate(nuint width, nuint height, nuint bitsPerComponent, nuint bitsPerPixel, nuint bytesPerRow,
            IntPtr space, uint bitmapInfo, IntPtr provider, IntPtr decode, [MarshalAs(UnmanagedType.I1)] bool shouldInterpolate, int intent);
        [DllImport(CG)] internal static extern nuint CGImageGetWidth(IntPtr image);
        [DllImport(CG)] internal static extern nuint CGImageGetHeight(IntPtr image);
        [DllImport(CG)] internal static extern IntPtr CGBitmapContextCreate(IntPtr data, nuint width, nuint height, nuint bitsPerComponent, nuint bytesPerRow, IntPtr space, uint bitmapInfo);
        [DllImport(CG)] internal static extern void CGContextSetBlendMode(IntPtr context, int mode);
        [DllImport(CG)] internal static extern void CGContextSetInterpolationQuality(IntPtr context, int quality);
        [DllImport(CG)] internal static extern void CGContextDrawImage(IntPtr context, Rect rect, IntPtr image);
        [DllImport(CG)] internal static extern void CGContextFlush(IntPtr context);
        [DllImport(IO)] internal static extern IntPtr CGImageDestinationCreateWithData(IntPtr data, IntPtr type, nuint count, IntPtr options);
        [DllImport(IO)] internal static extern void CGImageDestinationAddImage(IntPtr destination, IntPtr image, IntPtr properties);
        [DllImport(IO)] [return: MarshalAs(UnmanagedType.I1)] internal static extern bool CGImageDestinationFinalize(IntPtr destination);
        [DllImport(IO)] internal static extern IntPtr CGImageSourceCreateWithData(IntPtr data, IntPtr options);
        [DllImport(IO)] internal static extern IntPtr CGImageSourceCreateImageAtIndex(IntPtr source, nuint index, IntPtr options);
        [DllImport(IO)] internal static extern nuint CGImageSourceGetPrimaryImageIndex(IntPtr source);
        [DllImport(IO)] internal static extern IntPtr CGImageSourceCopyPropertiesAtIndex(IntPtr source, nuint index, IntPtr options);
        [DllImport(IO)] internal static extern IntPtr CGImageSourceCreateThumbnailAtIndex(IntPtr source, nuint index, IntPtr options);
    }
}
