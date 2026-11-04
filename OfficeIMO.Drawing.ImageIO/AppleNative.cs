using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Drawing.ImageIO;

/// <summary>Owned Core Foundation compatible object; Create/Copy results are released on every exit path.</summary>
internal sealed class AppleObject : SafeHandleZeroOrMinusOneIsInvalid {
    internal AppleObject(IntPtr value) : base(true) => SetHandle(value);
    internal IntPtr Value => handle;
    protected override bool ReleaseHandle() {
        AppleNative.CFRelease(handle);
        return true;
    }
}

/// <summary>Scoped access to exported framework constants, without retaining dlopen references.</summary>
internal sealed class AppleSymbols : IDisposable {
    private IntPtr _foundation;
    private IntPtr _imageIO;
    internal AppleSymbols() {
        _foundation = NativeLibrary.Load(AppleNative.Foundation);
        try { _imageIO = NativeLibrary.Load(AppleNative.ImageIO); }
        catch { Dispose(); throw; }
    }
    internal IntPtr Foundation(string name) => Marshal.ReadIntPtr(NativeLibrary.GetExport(_foundation, name));
    internal IntPtr ImageIO(string name) => Marshal.ReadIntPtr(NativeLibrary.GetExport(_imageIO, name));
    public void Dispose() {
        if (_imageIO != IntPtr.Zero) { NativeLibrary.Free(_imageIO); _imageIO = IntPtr.Zero; }
        if (_foundation != IntPtr.Zero) { NativeLibrary.Free(_foundation); _foundation = IntPtr.Zero; }
    }
}

/// <summary>64-bit Apple C ABI only; no AppKit/UIKit, Objective-C runtime or managed Apple workload.</summary>
internal static class AppleNative {
    internal const string Foundation = "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation";
    internal const string Graphics = "/System/Library/Frameworks/CoreGraphics.framework/CoreGraphics";
    internal const string ImageIO = "/System/Library/Frameworks/ImageIO.framework/ImageIO";
    internal const uint RgbaPremultipliedLast = 0x4001; // byteOrder32Big | premultipliedLast

    [StructLayout(LayoutKind.Sequential)]
    internal readonly struct Rect {
        internal readonly double X, Y, Width, Height;
        internal Rect(int width, int height) { X = 0; Y = 0; Width = width; Height = height; }
    }

    [DllImport(Foundation)] internal static extern void CFRelease(IntPtr value);
    [DllImport(Foundation)] internal static extern IntPtr CFDataCreate(IntPtr allocator, byte[] bytes, nint length);
    [DllImport(Foundation)] internal static extern IntPtr CFDictionaryCreate(IntPtr allocator, IntPtr[] keys, IntPtr[] values, nint count, IntPtr keyCallbacks, IntPtr valueCallbacks);
    [DllImport(Foundation)] internal static extern IntPtr CFDictionaryGetValue(IntPtr dictionary, IntPtr key);
    [DllImport(Foundation)] internal static extern nuint CFGetTypeID(IntPtr value);
    [DllImport(Foundation)] internal static extern nuint CFNumberGetTypeID();
    [DllImport(Foundation)] [return: MarshalAs(UnmanagedType.I1)]
    internal static extern bool CFNumberGetValue(IntPtr number, int type, out long value);
    [DllImport(ImageIO)] internal static extern IntPtr CGImageSourceCreateWithData(IntPtr data, IntPtr options);
    [DllImport(ImageIO)] internal static extern nuint CGImageSourceGetCount(IntPtr source);
    [DllImport(ImageIO)] internal static extern IntPtr CGImageSourceCopyPropertiesAtIndex(IntPtr source, nuint index, IntPtr options);
    [DllImport(ImageIO)] internal static extern IntPtr CGImageSourceCreateImageAtIndex(IntPtr source, nuint index, IntPtr options);
    [DllImport(Graphics)] internal static extern nuint CGImageGetWidth(IntPtr image);
    [DllImport(Graphics)] internal static extern nuint CGImageGetHeight(IntPtr image);
    [DllImport(Graphics)] internal static extern nuint CGImageGetBitsPerComponent(IntPtr image);
    [DllImport(Graphics)] internal static extern nuint CGImageGetBitsPerPixel(IntPtr image);
    [DllImport(Graphics)] internal static extern IntPtr CGImageGetColorSpace(IntPtr image);
    [DllImport(Graphics)] internal static extern int CGColorSpaceGetModel(IntPtr colorSpace);
    [DllImport(Graphics)] internal static extern IntPtr CGColorSpaceCreateDeviceRGB();
    [DllImport(Graphics)] internal static extern IntPtr CGColorSpaceCreateDeviceGray();
    [DllImport(Graphics)] internal static extern IntPtr CGImageCreateCopyWithColorSpace(IntPtr image, IntPtr space);
    [DllImport(Graphics)] internal static extern IntPtr CGBitmapContextCreate(IntPtr data, nuint width, nuint height, nuint bitsPerComponent, nuint bytesPerRow, IntPtr space, uint bitmapInfo);
    [DllImport(Graphics)] internal static extern void CGContextSetBlendMode(IntPtr context, int mode);
    [DllImport(Graphics)] internal static extern void CGContextSetInterpolationQuality(IntPtr context, int quality);
    [DllImport(Graphics)] internal static extern void CGContextDrawImage(IntPtr context, Rect rect, IntPtr image);
    [DllImport(Graphics)] internal static extern void CGContextFlush(IntPtr context);
}
