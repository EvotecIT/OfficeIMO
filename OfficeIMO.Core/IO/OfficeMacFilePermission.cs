using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace OfficeIMO.Core.Internal;

/// <summary>Retains the resolved native URL for the exact lifetime of a macOS file permission.</summary>
internal sealed class OfficeMacFilePermission : IDisposable {
    private const string Prefix = "officeimo.mac.v1:";
    private const int MaximumBookmarkBytes = 24_000;
    private const string CoreFoundation = "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation";
    private IntPtr _url;

    private OfficeMacFilePermission(IntPtr url) => _url = url;

    internal static bool IsNativeBookmark(string? bookmark) => bookmark?.StartsWith(Prefix, StringComparison.Ordinal) == true;

    internal static bool IsSandboxed {
        get {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return false;
            IntPtr task = SecTaskCreateFromSelf(IntPtr.Zero), value = IntPtr.Zero, error = IntPtr.Zero;
            IntPtr name = CFStringCreateWithCString(IntPtr.Zero, "com.apple.security.app-sandbox", 0x08000100);
            try {
                if (task == IntPtr.Zero || name == IntPtr.Zero) return false;
                value = SecTaskCopyValueForEntitlement(task, name, out error);
                return value != IntPtr.Zero && CFGetTypeID(value) == CFBooleanGetTypeID() && CFBooleanGetValue(value);
            } finally { Release(error); Release(value); Release(name); Release(task); }
        }
    }

    /// <summary>Creates an app-scoped reference while the picker-granted file is accessible.</summary>
    internal static string CreateBookmark(string path) {
        EnsureMac();
        byte[] name = Encoding.UTF8.GetBytes(Path.GetFullPath(path));
        IntPtr url = CFURLCreateFromFileSystemRepresentation(IntPtr.Zero, name, (IntPtr)name.Length, false);
        IntPtr data = IntPtr.Zero, error = IntPtr.Zero;
        try {
            if (url == IntPtr.Zero) throw new IOException("The selected file has no native URL.");
            data = CFURLCreateBookmarkData(IntPtr.Zero, url, (UIntPtr)(1u << 11), IntPtr.Zero, IntPtr.Zero, out error);
            if (data == IntPtr.Zero) throw new IOException("The selected file permission could not be saved.");
            long length = CFDataGetLength(data).ToInt64();
            if (length <= 0 || length > MaximumBookmarkBytes) throw new IOException("The native file permission exceeds the supported size.");
            var bytes = new byte[(int)length];
            Marshal.Copy(CFDataGetBytePtr(data), bytes, 0, bytes.Length);
            return Prefix + Convert.ToBase64String(bytes);
        } finally { Release(error); Release(data); Release(url); }
    }

    /// <summary>Resolves without UI or mounting volumes, verifies identity, then starts access.</summary>
    internal static OfficeMacFilePermission Open(string bookmark, string expectedPath) {
        EnsureMac();
        if (!IsNativeBookmark(bookmark) || bookmark.Length > Prefix.Length + MaximumBookmarkBytes * 4 / 3)
            throw new IOException("The native file permission is invalid or exceeds the supported size.");
        byte[] bytes;
        try { bytes = Convert.FromBase64String(bookmark.Substring(Prefix.Length)); }
        catch (FormatException exception) { throw new IOException("The native file permission is invalid.", exception); }
        if (bytes.Length == 0 || bytes.Length > MaximumBookmarkBytes) throw new IOException("The native file permission is empty or exceeds the supported size.");
        IntPtr data = CFDataCreate(IntPtr.Zero, bytes, (IntPtr)bytes.Length);
        IntPtr url = IntPtr.Zero, error = IntPtr.Zero;
        bool started = false, retained = false;
        try {
            if (data == IntPtr.Zero) throw new IOException("The native file permission could not be loaded.");
            url = CFURLCreateByResolvingBookmarkData(IntPtr.Zero, data,
                (UIntPtr)((1u << 8) | (1u << 9) | (1u << 10)), IntPtr.Zero, IntPtr.Zero, out _, out error);
            if (url == IntPtr.Zero) throw new IOException("The saved file permission is unavailable. Select the file again.");
            string resolved = CopyPath(url);
            if (!string.Equals(Path.GetFullPath(resolved), Path.GetFullPath(expectedPath), StringComparison.Ordinal))
                throw new IOException("The saved permission now identifies a different location. Select the file again.");
            started = CFURLStartAccessingSecurityScopedResource(url);
            if (!started) throw new UnauthorizedAccessException("The saved file permission has expired. Select the file again.");
            retained = true;
            return new OfficeMacFilePermission(url);
        } finally {
            if (!retained) { if (started) CFURLStopAccessingSecurityScopedResource(url); Release(url); }
            Release(error); Release(data);
        }
    }

    private static string CopyPath(IntPtr url) {
        IntPtr value = CFURLCopyFileSystemPath(url, IntPtr.Zero);
        try {
            if (value == IntPtr.Zero) throw new IOException("The saved permission is not a local file.");
            long characters = CFStringGetLength(value).ToInt64();
            if (characters <= 0 || characters > 4096) throw new IOException("The saved file path exceeds the supported size.");
            var bytes = new byte[checked((int)characters * 4 + 1)];
            if (!CFStringGetCString(value, bytes, (IntPtr)bytes.Length, 0x08000100))
                throw new IOException("The saved file path cannot be decoded.");
            return Encoding.UTF8.GetString(bytes, 0, Array.IndexOf(bytes, (byte)0));
        } finally { Release(value); }
    }

    public void Dispose() {
        IntPtr url = Interlocked.Exchange(ref _url, IntPtr.Zero);
        if (url == IntPtr.Zero) return;
        try { CFURLStopAccessingSecurityScopedResource(url); }
        finally { CFRelease(url); }
    }

    private static void EnsureMac() {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) throw new PlatformNotSupportedException("Native file permissions require macOS.");
    }
    private static void Release(IntPtr value) { if (value != IntPtr.Zero) CFRelease(value); }
    [DllImport(CoreFoundation)] private static extern IntPtr CFURLCreateFromFileSystemRepresentation(IntPtr allocator, byte[] bytes, IntPtr length, [MarshalAs(UnmanagedType.I1)] bool isDirectory);
    [DllImport(CoreFoundation)] private static extern IntPtr CFURLCreateBookmarkData(IntPtr allocator, IntPtr url, UIntPtr options, IntPtr properties, IntPtr relativeUrl, out IntPtr error);
    [DllImport(CoreFoundation)] private static extern IntPtr CFURLCreateByResolvingBookmarkData(IntPtr allocator, IntPtr data, UIntPtr options, IntPtr relativeUrl, IntPtr properties, out byte stale, out IntPtr error);
    [DllImport(CoreFoundation)] private static extern IntPtr CFURLCopyFileSystemPath(IntPtr url, IntPtr style);
    [DllImport(CoreFoundation)] private static extern IntPtr CFDataCreate(IntPtr allocator, byte[] bytes, IntPtr length);
    [DllImport(CoreFoundation)] private static extern IntPtr CFDataGetLength(IntPtr data);
    [DllImport(CoreFoundation)] private static extern IntPtr CFDataGetBytePtr(IntPtr data);
    [DllImport(CoreFoundation)] private static extern IntPtr CFStringGetLength(IntPtr value);
    [DllImport(CoreFoundation)] [return: MarshalAs(UnmanagedType.I1)] private static extern bool CFStringGetCString(IntPtr value, byte[] buffer, IntPtr size, uint encoding);
    [DllImport(CoreFoundation)] [return: MarshalAs(UnmanagedType.I1)] private static extern bool CFURLStartAccessingSecurityScopedResource(IntPtr url);
    [DllImport(CoreFoundation)] private static extern void CFURLStopAccessingSecurityScopedResource(IntPtr url);
    [DllImport(CoreFoundation)] private static extern void CFRelease(IntPtr value);
    // The only marshalled string is the ASCII entitlement key; LPStr also exists
    // in the netstandard2.0 contract. File paths use explicit UTF-8 byte buffers.
    [DllImport(CoreFoundation)] private static extern IntPtr CFStringCreateWithCString(IntPtr allocator, [MarshalAs(UnmanagedType.LPStr)] string text, uint encoding);
    [DllImport(CoreFoundation)] private static extern UIntPtr CFGetTypeID(IntPtr value);
    [DllImport(CoreFoundation)] private static extern UIntPtr CFBooleanGetTypeID();
    [DllImport(CoreFoundation)] [return: MarshalAs(UnmanagedType.I1)] private static extern bool CFBooleanGetValue(IntPtr value);
    [DllImport("/System/Library/Frameworks/Security.framework/Security")] private static extern IntPtr SecTaskCreateFromSelf(IntPtr allocator);
    [DllImport("/System/Library/Frameworks/Security.framework/Security")] private static extern IntPtr SecTaskCopyValueForEntitlement(IntPtr task, IntPtr name, out IntPtr error);
}
