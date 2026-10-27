using System;
using System.Runtime.InteropServices;

namespace OfficeIMO.Drawing.ImageIO;

/// <summary>Native sample extraction. No HEIF container parsing, composition or color policy lives here.</summary>
internal static class ImageIODecoder {
    internal static bool TryDecode(byte[] encoded, long maximumPixels, long maximumWorkingBytes, out OfficeRasterImage? result) {
        result = null;
        using var symbols = new AppleSymbols();
        IntPtr no = symbols.Foundation("kCFBooleanFalse");
        using var options = new AppleObject(AppleNative.CFDictionaryCreate(IntPtr.Zero,
            new[] { symbols.ImageIO("kCGImageSourceShouldCache"), symbols.ImageIO("kCGImageSourceShouldAllowFloat") },
            new[] { no, no }, 2, IntPtr.Zero, IntPtr.Zero));
        if (options.IsInvalid) return false;
        using var data = new AppleObject(AppleNative.CFDataCreate(IntPtr.Zero, encoded, encoded.Length));
        if (data.IsInvalid) return false;
        using var source = new AppleObject(AppleNative.CGImageSourceCreateWithData(data.Value, options.Value));
        if (source.IsInvalid || AppleNative.CGImageSourceGetCount(source.Value) == 0) return false;
        using var properties = new AppleObject(AppleNative.CGImageSourceCopyPropertiesAtIndex(source.Value, 0, options.Value));
        if (properties.IsInvalid ||
            !TryDimension(properties.Value, symbols.ImageIO("kCGImagePropertyPixelWidth"), out int width) ||
            !TryDimension(properties.Value, symbols.ImageIO("kCGImagePropertyPixelHeight"), out int height)) return false;
        long pixels = (long)width * height;
        // Account for input+CFData, up to 64-bit native samples, RGBA staging and FromRgba32's copy.
        // ImageIO's private decoder scratch is not exposed or controllable through this C API.
        if (pixels > maximumPixels || pixels * 16 + encoded.LongLength * 2 + 65_536 > maximumWorkingBytes) return false;
        using var decoded = new AppleObject(AppleNative.CGImageSourceCreateImageAtIndex(source.Value, 0, options.Value));
        if (decoded.IsInvalid || AppleNative.CGImageGetWidth(decoded.Value) != (nuint)width ||
            AppleNative.CGImageGetHeight(decoded.Value) != (nuint)height ||
            AppleNative.CGImageGetBitsPerComponent(decoded.Value) > 16 ||
            AppleNative.CGImageGetBitsPerPixel(decoded.Value) > 64) return false;
        IntPtr sourceSpace = AppleNative.CGImageGetColorSpace(decoded.Value); // borrowed
        if (sourceSpace == IntPtr.Zero) return false;
        int model = AppleNative.CGColorSpaceGetModel(sourceSpace);
        bool gray = model == 0;
        if (model != 1 && !gray || gray && AppleNative.CGImageGetBitsPerPixel(decoded.Value) > 32) return false;
        using var deviceSpace = new AppleObject(gray ? AppleNative.CGColorSpaceCreateDeviceGray() : AppleNative.CGColorSpaceCreateDeviceRGB());
        if (deviceSpace.IsInvalid) return false;
        // Retag the image, preserving component samples, before drawing. Drawing the original
        // source into an sRGB/deviceRGB destination would apply ICC/gamut conversion here.
        using var raw = new AppleObject(AppleNative.CGImageCreateCopyWithColorSpace(decoded.Value, deviceSpace.Value));
        if (raw.IsInvalid) return false;
        int components = gray ? 2 : 4;
        byte[] samples = new byte[checked(width * height * components)];
        GCHandle pin = GCHandle.Alloc(samples, GCHandleType.Pinned);
        try {
            using var context = new AppleObject(AppleNative.CGBitmapContextCreate(pin.AddrOfPinnedObject(),
                (nuint)width, (nuint)height, 8, (nuint)width * (nuint)components, deviceSpace.Value,
                gray ? 1U : AppleNative.RgbaPremultipliedLast)); // gray-alpha or RGBA, both premultiplied-last
            if (context.IsInvalid) return false;
            AppleNative.CGContextSetBlendMode(context.Value, 17); // kCGBlendModeCopy
            AppleNative.CGContextSetInterpolationQuality(context.Value, 1); // kCGInterpolationNone
            AppleNative.CGContextDrawImage(context.Value, new AppleNative.Rect(width, height), raw.Value);
            AppleNative.CGContextFlush(context.Value);
        } finally { pin.Free(); }
        byte[] rgba = samples;
        if (gray) {
            rgba = new byte[checked(width * height * 4)];
            for (int sourceOffset = 0, target = 0; sourceOffset < samples.Length; sourceOffset += 2, target += 4) {
                rgba[target] = rgba[target + 1] = rgba[target + 2] = samples[sourceOffset];
                rgba[target + 3] = samples[sourceOffset + 1];
            }
        }
        Unpremultiply(rgba);
        result = OfficeRasterImage.FromRgba32(width, height, rgba);
        return true;
    }

    private static bool TryDimension(IntPtr properties, IntPtr key, out int dimension) {
        dimension = 0;
        IntPtr number = AppleNative.CFDictionaryGetValue(properties, key);
        if (number == IntPtr.Zero || AppleNative.CFGetTypeID(number) != AppleNative.CFNumberGetTypeID() ||
            !AppleNative.CFNumberGetValue(number, 4, out long value) || value <= 0 || value > int.MaxValue) return false; // SInt64
        dimension = (int)value;
        return true;
    }

    private static void Unpremultiply(byte[] pixels) {
        for (int i = 0; i < pixels.Length; i += 4) {
            int alpha = pixels[i + 3];
            if (alpha == 255) continue;
            if (alpha == 0) { pixels[i] = pixels[i + 1] = pixels[i + 2] = 0; continue; }
            for (int channel = 0; channel < 3; channel++)
                pixels[i + channel] = (byte)Math.Min(255, (pixels[i + channel] * 255 + alpha / 2) / alpha);
        }
    }
}
