using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace OfficeIMO.Provenance;

/// <summary>Owns one bounded, private, read-only file snapshot used by multi-stage provenance operations.</summary>
internal sealed class OfficeProvenanceFileSnapshot : IDisposable {
    private const uint UnixOwnerDirectoryMode = 0x1C0; // 0700
    private const uint UnixOwnerReadOnlyMode = 0x100; // 0400
    private const uint UnixOwnerReadWriteMode = 0x180; // 0600
    private readonly string _directoryPath;
    private readonly FileStream _lease;
    private bool _leaseDisposed;
    private bool _disposed;

    private OfficeProvenanceFileSnapshot(string directoryPath, string filePath, long length, FileStream lease) {
        _directoryPath = directoryPath;
        FilePath = filePath;
        Length = length;
        _lease = lease;
    }

    /// <summary>Gets the immutable snapshot path supplied to format owners and assessment providers.</summary>
    internal string FilePath { get; }

    /// <summary>Gets the captured encoded byte length.</summary>
    internal long Length { get; }

    /// <summary>Captures one input through a single bounded read and holds a shared-read lease until disposal.</summary>
    internal static OfficeProvenanceFileSnapshot Capture(
        string sourcePath,
        long maximumBytes,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(sourcePath)) throw new ArgumentException("A source path is required.", nameof(sourcePath));
        if (maximumBytes <= 0 || maximumBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(maximumBytes));

        string fullPath = Path.GetFullPath(sourcePath);
        string directoryPath = CreatePrivateDirectory();
        string extension = Path.GetExtension(fullPath);
        string filePath = Path.Combine(directoryPath, "snapshot" + extension);
        try {
            long copiedBytes = 0;
            using (var source = new FileStream(
                       fullPath,
                       FileMode.Open,
                       FileAccess.Read,
                       FileShare.Read,
                       81920,
                       FileOptions.SequentialScan))
            using (var destination = new FileStream(
                       filePath,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.None,
                       81920,
                       FileOptions.SequentialScan)) {
                if (source.Length > maximumBytes) {
                    throw new InvalidDataException("The asset snapshot exceeds the configured input limit.");
                }

                var buffer = new byte[81920];
                int read;
                while ((read = source.Read(buffer, 0, buffer.Length)) != 0) {
                    cancellationToken.ThrowIfCancellationRequested();
                    copiedBytes += read;
                    if (copiedBytes > maximumBytes) {
                        throw new InvalidDataException("The asset snapshot exceeds the configured input limit.");
                    }
                    destination.Write(buffer, 0, read);
                }
                cancellationToken.ThrowIfCancellationRequested();
                destination.Flush(flushToDisk: true);
            }

            MakeReadOnly(filePath);
            var lease = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                1,
                FileOptions.RandomAccess);
            return new OfficeProvenanceFileSnapshot(directoryPath, filePath, copiedBytes, lease);
        } catch {
            Cleanup(filePath, directoryPath, throwIfSensitiveDataRemains: true);
            throw;
        }
    }

    /// <inheritdoc />
    public void Dispose() {
        if (_disposed) return;
        if (!_leaseDisposed) {
            _lease.Dispose();
            _leaseDisposed = true;
        }
        Cleanup(FilePath, _directoryPath, throwIfSensitiveDataRemains: true);
        _disposed = true;
    }

    private static string CreatePrivateDirectory() {
        string tempPath = Path.GetTempPath();
        for (int attempt = 0; attempt < 16; attempt++) {
            string path = Path.Combine(tempPath, "officeimo-provenance-" + Guid.NewGuid().ToString("N"));
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                Directory.CreateDirectory(path);
                return path;
            }
            if (UnixMkdir(path, UnixOwnerDirectoryMode) == 0) return path;
            int error = Marshal.GetLastWin32Error();
            if (error != 17) throw new IOException($"Unable to create a private provenance snapshot directory (errno {error}).");
        }
        throw new IOException("Unable to allocate a unique private provenance snapshot directory.");
    }

    private static void MakeReadOnly(string path) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            File.SetAttributes(path, FileAttributes.ReadOnly);
            return;
        }
        if (UnixChmod(path, UnixOwnerReadOnlyMode) != 0) {
            throw new IOException($"Unable to protect the provenance snapshot (errno {Marshal.GetLastWin32Error()}).");
        }
    }

    private static void MakeWritable(string path) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            File.SetAttributes(path, FileAttributes.Normal);
            return;
        }
        if (UnixChmod(path, UnixOwnerReadWriteMode) != 0) {
            throw new IOException($"Unable to unlock the provenance snapshot for cleanup (errno {Marshal.GetLastWin32Error()}).");
        }
    }

    private static void Cleanup(string filePath, string directoryPath, bool throwIfSensitiveDataRemains) {
        Exception? cleanupFailure = null;
        if (File.Exists(filePath)) {
            try {
                MakeWritable(filePath);
                File.Delete(filePath);
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                cleanupFailure = exception;
                try {
                    using var scrub = new FileStream(filePath, FileMode.Truncate, FileAccess.Write, FileShare.Read);
                    scrub.Flush(flushToDisk: true);
                    try { File.Delete(filePath); } catch (Exception deleteException) when (deleteException is IOException or UnauthorizedAccessException) { }
                    cleanupFailure = null;
                } catch (Exception scrubException) when (scrubException is IOException or UnauthorizedAccessException) {
                    cleanupFailure = scrubException;
                }
            }
        }

        try {
            if (Directory.Exists(directoryPath)) Directory.Delete(directoryPath, recursive: false);
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            if (File.Exists(filePath)) cleanupFailure ??= exception;
        }

        if (throwIfSensitiveDataRemains && cleanupFailure != null && File.Exists(filePath)) {
            throw new IOException(
                $"The provenance snapshot could not be removed or erased; '{filePath}' is retained for operator cleanup.",
                cleanupFailure);
        }
    }

    [DllImport("libc", SetLastError = true, EntryPoint = "mkdir")]
    private static extern int UnixMkdir(string path, uint mode);

    [DllImport("libc", SetLastError = true, EntryPoint = "chmod")]
    private static extern int UnixChmod(string path, uint mode);
}
