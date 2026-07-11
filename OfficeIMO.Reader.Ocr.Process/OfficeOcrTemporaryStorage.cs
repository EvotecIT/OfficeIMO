using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;

namespace OfficeIMO.Reader.Ocr.Process;

/// <summary>Creates private per-request storage for OCR payloads and provider output.</summary>
internal static class OfficeOcrTemporaryStorage {
    private const uint OwnerDirectoryMode = 0x1C0; // 0700
    private const uint OwnerFileMode = 0x180; // 0600

    internal static string CreateRequestDirectory(string temporaryRoot, string prefix) {
        Directory.CreateDirectory(temporaryRoot);
        string requestDirectory = Path.Combine(temporaryRoot, prefix + Guid.NewGuid().ToString("N"));
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            Directory.CreateDirectory(requestDirectory);
            return requestDirectory;
        }

        if (CreateDirectory(requestDirectory, OwnerDirectoryMode) != 0) {
            throw BuildIOException("create the private OCR request directory", Marshal.GetLastWin32Error());
        }
        if (ChangeMode(requestDirectory, OwnerDirectoryMode) == 0) return requestDirectory;

        int error = Marshal.GetLastWin32Error();
        TryDeleteDirectory(requestDirectory);
        throw BuildIOException("set private permissions on the OCR request directory", error);
    }

    internal static void WriteAllBytes(string path, byte[] payload) {
        if (payload == null) throw new ArgumentNullException(nameof(payload));
        using FileStream stream = CreatePrivateFile(path);
        stream.Write(payload, 0, payload.Length);
    }

    internal static void EnsurePrivateFile(string path) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        if (ChangeMode(path, OwnerFileMode) != 0) {
            throw BuildIOException("set private permissions on an OCR temporary file", Marshal.GetLastWin32Error());
        }
    }

    private static FileStream CreatePrivateFile(string path) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            return new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        }

        int descriptor = CreateFile(path, OwnerFileMode);
        if (descriptor < 0) throw BuildIOException("create a private OCR temporary file", Marshal.GetLastWin32Error());
        var handle = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
        try {
            return new FileStream(handle, FileAccess.Write);
        } catch {
            handle.Dispose();
            throw;
        }
    }

    private static IOException BuildIOException(string operation, int error) {
        return new IOException("Unable to " + operation + " (OS error " + error + ").");
    }

    private static void TryDeleteDirectory(string path) {
        try {
            if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
        } catch (IOException) {
        } catch (UnauthorizedAccessException) {
        }
    }

    [DllImport("libc", EntryPoint = "mkdir", SetLastError = true)]
    private static extern int CreateDirectory(string path, uint mode);

    [DllImport("libc", EntryPoint = "creat", SetLastError = true)]
    private static extern int CreateFile(string path, uint mode);

    [DllImport("libc", EntryPoint = "chmod", SetLastError = true)]
    private static extern int ChangeMode(string path, uint mode);
}
