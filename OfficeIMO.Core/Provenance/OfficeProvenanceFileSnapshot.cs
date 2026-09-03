using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using OfficeIMO.Internal;

namespace OfficeIMO.Provenance;

/// <summary>Owns one bounded, private, read-only file snapshot used by multi-stage provenance operations.</summary>
internal sealed class OfficeProvenanceFileSnapshot : IDisposable {
    private const uint UnixOwnerDirectoryMode = 0x1C0; // 0700
    private const uint UnixOwnerReadOnlyMode = 0x100; // 0400
    private const uint UnixOwnerReadWriteMode = 0x180; // 0600
    private readonly string _directoryPath;
    private readonly FileStream _lease;
    private readonly string _physicalIdentity;
    private readonly byte[] _sha256;
    private readonly List<string> _dependentFiles = new List<string>();
    private readonly List<DependencySnapshot> _dependentSnapshots = new List<DependencySnapshot>();
    private bool _dependentLeasesDisposed;
    private bool _leaseDisposed;
    private bool _disposed;

    private OfficeProvenanceFileSnapshot(
        string directoryPath,
        string filePath,
        long length,
        string physicalIdentity,
        byte[] sha256,
        FileStream lease) {
        _directoryPath = directoryPath;
        FilePath = filePath;
        Length = length;
        _physicalIdentity = physicalIdentity;
        _sha256 = sha256;
        _lease = lease;
    }

    /// <summary>Gets the immutable snapshot path supplied to format owners and assessment providers.</summary>
    internal string FilePath { get; }

    /// <summary>Gets the captured encoded byte length.</summary>
    internal long Length { get; }

    /// <summary>Captures local relative manifest references beside the immutable asset snapshot.</summary>
    internal void CaptureExternalManifestDependencies(
        string sourcePath,
        OfficeProvenanceReport report,
        long maximumDependencyBytes,
        long maximumTotalBytes,
        CancellationToken cancellationToken = default) {
        if (_disposed) throw new ObjectDisposedException(nameof(OfficeProvenanceFileSnapshot));
        if (string.IsNullOrWhiteSpace(sourcePath)) throw new ArgumentException("A source path is required.", nameof(sourcePath));
        if (report == null) throw new ArgumentNullException(nameof(report));
        if (maximumDependencyBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumDependencyBytes));
        if (maximumTotalBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumTotalBytes));

        string sourceDirectory = Path.GetDirectoryName(Path.GetFullPath(sourcePath))!;
        string physicalSourceDirectory = OfficePathIdentity.ResolvePhysicalPath(sourceDirectory);
        long capturedBytes = 0;
        var capturedTargets = new HashSet<string>(OfficePathIdentity.GetComparer(_directoryPath));
        foreach (OfficeProvenanceEvidence evidence in report.Evidence.Where(item =>
                     item.IsStructurallyValid &&
                     item.Carrier == OfficeProvenanceCarrierKind.C2paExternalManifest &&
                     !string.IsNullOrWhiteSpace(item.Value))) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryResolveRelativeDependency(
                    sourceDirectory,
                    _directoryPath,
                    FilePath,
                    report.GetExternalManifestReference(evidence)!,
                    out string sourceDependency,
                    out string targetDependency) ||
                !capturedTargets.Add(targetDependency) ||
                !File.Exists(sourceDependency)) continue;

            string? targetDirectory = Path.GetDirectoryName(targetDependency);
            if (!string.IsNullOrEmpty(targetDirectory)) Directory.CreateDirectory(targetDirectory);
            _dependentFiles.Add(targetDependency);
            try {
                long copiedBytes = CopyDependency(
                    sourceDependency,
                    targetDependency,
                    physicalSourceDirectory,
                    maximumDependencyBytes,
                    maximumTotalBytes - report.ExpandedInspectionBytes - capturedBytes,
                    cancellationToken);
                capturedBytes += copiedBytes;
                MakeReadOnly(targetDependency);
                _dependentSnapshots.Add(DependencySnapshot.Capture(targetDependency, cancellationToken));
            } catch {
                _ = TryEraseAndDeleteFile(targetDependency);
                throw;
            }
        }
    }

    /// <summary>Verifies that captured external manifests stayed identical throughout provider execution.</summary>
    internal void VerifyExternalManifestDependencies(CancellationToken cancellationToken = default) {
        if (_disposed) throw new ObjectDisposedException(nameof(OfficeProvenanceFileSnapshot));
        foreach (DependencySnapshot dependency in _dependentSnapshots) {
            dependency.Verify(cancellationToken);
        }
    }

    /// <summary>Verifies that the primary snapshot stayed identical throughout owner and provider execution.</summary>
    internal void VerifyPrimaryFile(CancellationToken cancellationToken = default) {
        if (_disposed) throw new ObjectDisposedException(nameof(OfficeProvenanceFileSnapshot));
        try {
            string physicalDirectory = OfficePathIdentity.ResolvePhysicalPath(_directoryPath);
            using FileStream stream = OfficePathIdentity.OpenRegularFileForRead(
                FilePath,
                physicalDirectory,
                81920);
            string identity = OfficePathIdentity.GetPhysicalIdentityKey(FilePath, stream.SafeFileHandle);
            if (stream.Length != Length ||
                !string.Equals(identity, _physicalIdentity, StringComparison.Ordinal) ||
                !FixedTimeEquals(ComputeHash(stream, cancellationToken), _sha256)) {
                throw new InvalidDataException(
                    "The primary provenance snapshot changed while the operation was running.");
            }
        } catch (InvalidDataException) {
            throw;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new InvalidDataException(
                "The primary provenance snapshot disappeared or became inaccessible while the operation was running.",
                exception);
        }
    }

    /// <summary>Captures one input through a single bounded read and holds a shared-read lease until disposal.</summary>
    internal static OfficeProvenanceFileSnapshot Capture(
        string sourcePath,
        long maximumBytes,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(sourcePath)) throw new ArgumentException("A source path is required.", nameof(sourcePath));
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));

        string fullPath = Path.GetFullPath(sourcePath);
        string directoryPath = CreatePrivateDirectory();
        string filePath = Path.Combine(directoryPath, Path.GetFileName(fullPath));
        try {
            long copiedBytes = 0;
            string physicalSourceDirectory = OfficePathIdentity.ResolvePhysicalPath(
                Path.GetDirectoryName(fullPath)!);
            using (FileStream source = OpenSnapshotSource(fullPath, physicalSourceDirectory))
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
                81920,
                FileOptions.SequentialScan);
            try {
                string physicalIdentity = OfficePathIdentity.GetPhysicalIdentityKey(filePath, lease.SafeFileHandle);
                byte[] sha256 = ComputeHash(lease, cancellationToken);
                lease.Position = 0;
                return new OfficeProvenanceFileSnapshot(
                    directoryPath,
                    filePath,
                    copiedBytes,
                    physicalIdentity,
                    sha256,
                    lease);
            } catch {
                lease.Dispose();
                throw;
            }
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
        if (!_dependentLeasesDisposed) {
            foreach (DependencySnapshot dependency in _dependentSnapshots) dependency.Dispose();
            _dependentLeasesDisposed = true;
        }
        Cleanup(FilePath, _directoryPath, throwIfSensitiveDataRemains: true, _dependentFiles);
        _disposed = true;
    }

    private static long CopyDependency(
        string sourcePath,
        string targetPath,
        string physicalSourceDirectory,
        long maximumDependencyBytes,
        long remainingTotalBytes,
        CancellationToken cancellationToken) {
        if (remainingTotalBytes <= 0) {
            throw new InvalidDataException("External provenance manifests exceed the configured expanded-data limit.");
        }
        long copiedBytes = 0;
        FileStream source;
        try {
            source = OfficePathIdentity.OpenRegularFileForRead(
                sourcePath,
                physicalSourceDirectory,
                81920);
        } catch (InvalidDataException) {
            throw;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new InvalidDataException(
                "The external provenance manifest could not be opened as a regular file within the source directory.",
                exception);
        }
        using (source) {
            if (source.Length > maximumDependencyBytes) {
                throw new InvalidDataException("An external provenance manifest exceeds the configured manifest limit.");
            }
            if (source.Length > remainingTotalBytes) {
                throw new InvalidDataException("External provenance manifests exceed the configured expanded-data limit.");
            }
            using var target = new FileStream(targetPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, FileOptions.SequentialScan);
            var buffer = new byte[81920];
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) != 0) {
                cancellationToken.ThrowIfCancellationRequested();
                copiedBytes += read;
                if (copiedBytes > maximumDependencyBytes) {
                    throw new InvalidDataException("An external provenance manifest exceeds the configured manifest limit.");
                }
                if (copiedBytes > remainingTotalBytes) {
                    throw new InvalidDataException("External provenance manifests exceed the configured expanded-data limit.");
                }
                target.Write(buffer, 0, read);
            }
            target.Flush(flushToDisk: true);
            return copiedBytes;
        }
    }

    private static bool TryResolveRelativeDependency(
        string sourceDirectory,
        string snapshotDirectory,
        string snapshotFilePath,
        string reference,
        out string sourcePath,
        out string targetPath) {
        sourcePath = string.Empty;
        targetPath = string.Empty;
        if (!Uri.TryCreate(reference, UriKind.RelativeOrAbsolute, out Uri? uri) || uri.IsAbsoluteUri) return false;
        string relative = reference;
        int delimiter = relative.IndexOfAny(new[] { '?', '#' });
        if (delimiter >= 0) relative = relative.Substring(0, delimiter);
        if (relative.Length == 0) return false;
        try {
            relative = Uri.UnescapeDataString(relative).Replace('/', Path.DirectorySeparatorChar);
            if (Path.IsPathRooted(relative)) return false;
            sourcePath = Path.GetFullPath(Path.Combine(sourceDirectory, relative));
            targetPath = Path.GetFullPath(Path.Combine(snapshotDirectory, relative));
            return IsWithinDirectory(sourcePath, sourceDirectory) &&
                   IsWithinDirectory(targetPath, snapshotDirectory) &&
                   !PathsEqual(targetPath, snapshotFilePath);
        } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or UriFormatException) {
            sourcePath = string.Empty;
            targetPath = string.Empty;
            return false;
        }
    }

    private static bool IsWithinDirectory(string path, string directory) {
        string root = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        StringComparison comparison = OfficePathIdentity.GetComparison(directory);
        return path.StartsWith(root, comparison);
    }

    private static bool PathsEqual(string left, string right) => string.Equals(
        left,
        right,
        OfficePathIdentity.GetComparison(left));

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

    private static void Cleanup(
        string filePath,
        string directoryPath,
        bool throwIfSensitiveDataRemains,
        IEnumerable<string>? dependentFiles = null) {
        Exception? cleanupFailure = null;
        string[] dependencies = (dependentFiles ?? Array.Empty<string>()).ToArray();
        foreach (string candidate in dependencies.AsEnumerable().Reverse().Concat(new[] { filePath })) {
            cleanupFailure = TryEraseAndDeleteFile(candidate) ?? cleanupFailure;
        }

        foreach (string? directory in dependencies
                     .SelectMany(path => EnumerateDependencyDirectories(path, directoryPath))
                     .Where(path => !string.IsNullOrEmpty(path) && !PathsEqual(path!, directoryPath))
                     .Distinct(OfficePathIdentity.GetComparer(directoryPath))
                     .OrderByDescending(path => path!.Length)) {
            try {
                if (Directory.Exists(directory)) Directory.Delete(directory!, recursive: false);
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                cleanupFailure ??= exception;
            }
        }

        try {
            if (Directory.Exists(directoryPath)) Directory.Delete(directoryPath, recursive: false);
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            cleanupFailure ??= exception;
        }

        if (throwIfSensitiveDataRemains && cleanupFailure != null && Directory.Exists(directoryPath)) {
            string retainedPath = File.Exists(filePath) ? filePath : directoryPath;
            throw new IOException(
                $"The provenance snapshot could not be removed or erased; '{retainedPath}' is retained for operator cleanup.",
                cleanupFailure);
        }
    }

    private static IEnumerable<string> EnumerateDependencyDirectories(string filePath, string rootDirectory) {
        string? current = Path.GetDirectoryName(filePath);
        while (!string.IsNullOrEmpty(current) &&
               !PathsEqual(current, rootDirectory) &&
               IsWithinDirectory(current, rootDirectory)) {
            yield return current;
            current = Path.GetDirectoryName(current);
        }
    }

    private sealed class DependencySnapshot : IDisposable {
        private readonly string _path;
        private readonly long _length;
        private readonly byte[] _sha256;
        private readonly FileStream _lease;

        private DependencySnapshot(string path, long length, byte[] sha256, FileStream lease) {
            _path = path;
            _length = length;
            _sha256 = sha256;
            _lease = lease;
        }

        internal static DependencySnapshot Capture(string path, CancellationToken cancellationToken) {
            var lease = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                81920,
                FileOptions.SequentialScan);
            try {
                byte[] sha256 = OfficeProvenanceFileSnapshot.ComputeHash(lease, cancellationToken);
                lease.Position = 0;
                return new DependencySnapshot(path, lease.Length, sha256, lease);
            } catch {
                lease.Dispose();
                throw;
            }
        }

        internal void Verify(CancellationToken cancellationToken) {
            try {
                using var stream = new FileStream(
                    _path,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read,
                    81920,
                    FileOptions.SequentialScan);
                if (stream.Length != _length ||
                    !OfficeProvenanceFileSnapshot.FixedTimeEquals(
                        OfficeProvenanceFileSnapshot.ComputeHash(stream, cancellationToken),
                        _sha256)) {
                    throw new InvalidDataException(
                        "An external provenance manifest changed while the assessment was running.");
                }
            } catch (Exception exception) when (exception is FileNotFoundException or DirectoryNotFoundException) {
                throw new InvalidDataException(
                    "An external provenance manifest disappeared while the assessment was running.",
                    exception);
            }
        }

        public void Dispose() => _lease.Dispose();
    }

    private static FileStream OpenSnapshotSource(string path, string physicalSourceDirectory) {
        try {
            return OfficePathIdentity.OpenRegularFileForRead(path, physicalSourceDirectory, 81920);
        } catch (InvalidDataException) {
            throw;
        } catch (FileNotFoundException) {
            throw;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new InvalidDataException(
                "The provenance snapshot source could not be opened as a regular file.",
                exception);
        }
    }

    private static byte[] ComputeHash(Stream stream, CancellationToken cancellationToken) {
        using IncrementalHash algorithm = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[81920];
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) != 0) {
            cancellationToken.ThrowIfCancellationRequested();
            algorithm.AppendData(buffer, 0, read);
        }
        return algorithm.GetHashAndReset();
    }

    private static bool FixedTimeEquals(byte[] left, byte[] right) {
        if (left.Length != right.Length) return false;
        int difference = 0;
        for (int index = 0; index < left.Length; index++) difference |= left[index] ^ right[index];
        return difference == 0;
    }

    private static Exception? TryEraseAndDeleteFile(string path) {
        if (!File.Exists(path)) return null;
        try {
            MakeWritable(path);
            File.Delete(path);
            return null;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            try {
                using var scrub = new FileStream(path, FileMode.Truncate, FileAccess.Write, FileShare.Read);
                scrub.Flush(flushToDisk: true);
                try { File.Delete(path); } catch (Exception deleteException) when (deleteException is IOException or UnauthorizedAccessException) { }
                return File.Exists(path) ? exception : null;
            } catch (Exception scrubException) when (scrubException is IOException or UnauthorizedAccessException) {
                return scrubException;
            }
        }
    }

    [DllImport("libc", SetLastError = true, EntryPoint = "mkdir")]
    private static extern int UnixMkdir(string path, uint mode);

    [DllImport("libc", SetLastError = true, EntryPoint = "chmod")]
    private static extern int UnixChmod(string path, uint mode);
}
