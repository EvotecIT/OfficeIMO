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
    private const uint UnixOwnerReadExecuteDirectoryMode = 0x140; // 0500
    private const uint UnixOwnerReadOnlyMode = 0x100; // 0400
    private const uint UnixOwnerReadWriteMode = 0x180; // 0600
    private readonly string _directoryPath;
    private readonly FileStream _lease;
    private readonly string? _physicalIdentity;
    private readonly bool _usesPhysicalIdentity;
    private readonly byte[] _sha256;
    private readonly List<string> _dependentFiles = new List<string>();
    private readonly List<DependencySnapshot> _dependentSnapshots = new List<DependencySnapshot>();
    private bool _dependentLeasesDisposed;
    private bool _leaseDisposed;
    private bool _directorySealed;
    private bool _disposed;

    private OfficeProvenanceFileSnapshot(
        string directoryPath,
        string filePath,
        long length,
        string? physicalIdentity,
        byte[] sha256,
        FileStream lease,
        bool usesPhysicalIdentity) {
        _directoryPath = directoryPath;
        FilePath = filePath;
        Length = length;
        _physicalIdentity = physicalIdentity;
        _sha256 = sha256;
        _lease = lease;
        _usesPhysicalIdentity = usesPhysicalIdentity;
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

        bool resealDirectory = _directorySealed;
        if (resealDirectory) MakeSnapshotDirectoriesWritable();
        try {
            string sourceDirectory = Path.GetDirectoryName(Path.GetFullPath(sourcePath))!;
            string sourceDirectoryIdentity = _usesPhysicalIdentity
                ? OfficePathIdentity.ResolvePhysicalPath(sourceDirectory)
                : Path.GetFullPath(sourceDirectory);
            long capturedBytes = 0;
            var capturedTargets = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeProvenanceEvidence evidence in report.Evidence.Where(item =>
                         item.Carrier == OfficeProvenanceCarrierKind.C2paExternalManifest &&
                         !string.IsNullOrWhiteSpace(item.Value))) {
                cancellationToken.ThrowIfCancellationRequested();
                string reference = report.GetExternalManifestReference(evidence)!;
                if (Uri.TryCreate(reference, UriKind.Absolute, out Uri? absoluteReference)) {
                    if (absoluteReference.IsFile) {
                        throw new InvalidDataException(
                            "An absolute file-based external provenance manifest cannot be bound to the immutable snapshot.");
                    }
                    continue;
                }
                if (!evidence.IsStructurallyValid) continue;
                if (!TryResolveRelativeDependency(
                        sourceDirectory,
                        _directoryPath,
                        FilePath,
                        reference,
                        out string sourceDependency,
                        out string targetDependency)) {
                    throw new InvalidDataException(
                        "A relative external provenance manifest cannot be bound within the immutable snapshot.");
                }
                if (!capturedTargets.Add(GetDependencyTargetIdentity(targetDependency)) ||
                    !File.Exists(sourceDependency)) continue;

                string? targetDirectory = Path.GetDirectoryName(targetDependency);
                if (!string.IsNullOrEmpty(targetDirectory)) Directory.CreateDirectory(targetDirectory);
                _dependentFiles.Add(targetDependency);
                try {
                    long copiedBytes = CopyDependency(
                        sourceDependency,
                        targetDependency,
                        sourceDirectoryIdentity,
                        _usesPhysicalIdentity,
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
        } finally {
            if (resealDirectory) SealForProviderAccess();
        }
    }

    /// <summary>Removes directory write access while path-based owners and providers consume the snapshot.</summary>
    internal void SealForProviderAccess() {
        if (_disposed) throw new ObjectDisposedException(nameof(OfficeProvenanceFileSnapshot));
        if (_directorySealed) return;
        _directorySealed = true;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        foreach (string directory in EnumerateSnapshotDirectories().OrderByDescending(path => path.Length)) {
            if (Directory.Exists(directory) && UnixChmod(directory, UnixOwnerReadExecuteDirectoryMode) != 0) {
                throw new IOException(
                    $"Unable to seal the provenance snapshot directory (errno {Marshal.GetLastWin32Error()}).");
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

    private string GetDependencyTargetIdentity(string targetPath) => _usesPhysicalIdentity
        ? OfficePathIdentity.Normalize(targetPath)
        : Path.GetFullPath(targetPath);

    /// <summary>Verifies that the primary snapshot stayed identical throughout owner and provider execution.</summary>
    internal void VerifyPrimaryFile(CancellationToken cancellationToken = default) {
        if (_disposed) throw new ObjectDisposedException(nameof(OfficeProvenanceFileSnapshot));
        try {
            string directoryIdentity = _usesPhysicalIdentity
                ? OfficePathIdentity.ResolvePhysicalPath(_directoryPath)
                : Path.GetFullPath(_directoryPath);
            using FileStream stream = OpenRegularFileForRead(
                FilePath,
                directoryIdentity,
                _usesPhysicalIdentity,
                "The primary provenance snapshot could not be opened as a regular file.");
            string? identity = _usesPhysicalIdentity
                ? OfficePathIdentity.GetPhysicalIdentityKey(FilePath, stream.SafeFileHandle)
                : null;
            if (stream.Length != Length ||
                (_usesPhysicalIdentity && !string.Equals(identity, _physicalIdentity, StringComparison.Ordinal)) ||
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
        CancellationToken cancellationToken = default) => CaptureCore(
            sourcePath,
            maximumBytes,
            cancellationToken,
            OfficePathIdentity.SupportsPhysicalIdentity);

    /// <summary>Exercises the portable snapshot path on platforms where physical file identity is unavailable.</summary>
    internal static OfficeProvenanceFileSnapshot CapturePortable(
        string sourcePath,
        long maximumBytes,
        CancellationToken cancellationToken = default) => CaptureCore(
            sourcePath,
            maximumBytes,
            cancellationToken,
            usesPhysicalIdentity: false);

    private static OfficeProvenanceFileSnapshot CaptureCore(
        string sourcePath,
        long maximumBytes,
        CancellationToken cancellationToken,
        bool usesPhysicalIdentity) {
        if (string.IsNullOrWhiteSpace(sourcePath)) throw new ArgumentException("A source path is required.", nameof(sourcePath));
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));

        string fullPath = Path.GetFullPath(sourcePath);
        string directoryPath = CreatePrivateDirectory();
        string filePath = Path.Combine(directoryPath, Path.GetFileName(fullPath));
        try {
            long copiedBytes;
            byte[] copiedSha256;
            string sourceDirectory = Path.GetDirectoryName(fullPath)!;
            string sourceDirectoryIdentity = usesPhysicalIdentity
                ? OfficePathIdentity.ResolvePhysicalPath(sourceDirectory)
                : Path.GetFullPath(sourceDirectory);
            using (FileStream source = OpenSnapshotSource(fullPath, sourceDirectoryIdentity, usesPhysicalIdentity))
            using (var destination = new FileStream(
                       filePath,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.None,
                       81920,
                       FileOptions.SequentialScan)) {
                copiedSha256 = CopyStableSource(
                    source,
                    destination,
                    maximumBytes,
                    "The asset snapshot exceeds the configured input limit.",
                    cancellationToken,
                    out copiedBytes);
                if (usesPhysicalIdentity && !OfficePathIdentity.IsOpenedFileWithinRootByIdentity(
                        fullPath,
                        sourceDirectoryIdentity,
                        source.SafeFileHandle)) {
                    throw new InvalidDataException(
                        "The asset snapshot source changed identity while it was being captured.");
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
                string? physicalIdentity = usesPhysicalIdentity
                    ? OfficePathIdentity.GetPhysicalIdentityKey(filePath, lease.SafeFileHandle)
                    : null;
                byte[] sha256 = ComputeHash(lease, cancellationToken);
                if (lease.Length != copiedBytes || !FixedTimeEquals(sha256, copiedSha256)) {
                    throw new InvalidDataException(
                        "The primary provenance snapshot did not match the stable source bytes that were captured.");
                }
                lease.Position = 0;
                return new OfficeProvenanceFileSnapshot(
                    directoryPath,
                    filePath,
                    copiedBytes,
                    physicalIdentity,
                    sha256,
                    lease,
                    usesPhysicalIdentity);
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
        MakeSnapshotDirectoriesWritable();
        Cleanup(FilePath, _directoryPath, throwIfSensitiveDataRemains: true, _dependentFiles);
        _disposed = true;
    }

    private void MakeSnapshotDirectoriesWritable() {
        if (!_directorySealed) return;
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            foreach (string directory in EnumerateSnapshotDirectories().OrderBy(path => path.Length)) {
                if (Directory.Exists(directory) && UnixChmod(directory, UnixOwnerDirectoryMode) != 0) {
                    throw new IOException(
                        $"Unable to unlock the provenance snapshot directory for cleanup (errno {Marshal.GetLastWin32Error()}).");
                }
            }
        }
        _directorySealed = false;
    }

    private IEnumerable<string> EnumerateSnapshotDirectories() => _dependentFiles
        .SelectMany(path => EnumerateDependencyDirectories(path, _directoryPath))
        .Concat(new[] { _directoryPath })
        .Distinct(OfficePathIdentity.GetComparer(_directoryPath));

    private static long CopyDependency(
        string sourcePath,
        string targetPath,
        string sourceDirectoryIdentity,
        bool usesPhysicalIdentity,
        long maximumDependencyBytes,
        long remainingTotalBytes,
        CancellationToken cancellationToken) {
        if (remainingTotalBytes <= 0) {
            throw new InvalidDataException("External provenance manifests exceed the configured expanded-data limit.");
        }
        long copiedBytes = 0;
        FileStream source = OpenRegularFileForRead(
            sourcePath,
            sourceDirectoryIdentity,
            usesPhysicalIdentity,
            "The external provenance manifest could not be opened as a regular file within the source directory.");
        using (source) {
            if (source.Length > maximumDependencyBytes) {
                throw new InvalidDataException("An external provenance manifest exceeds the configured manifest limit.");
            }
            if (source.Length > remainingTotalBytes) {
                throw new InvalidDataException("External provenance manifests exceed the configured expanded-data limit.");
            }
            using var target = new FileStream(targetPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, FileOptions.SequentialScan);
            long effectiveMaximum = Math.Min(maximumDependencyBytes, remainingTotalBytes);
            string limitMessage = remainingTotalBytes < maximumDependencyBytes
                ? "External provenance manifests exceed the configured expanded-data limit."
                : "An external provenance manifest exceeds the configured manifest limit.";
            _ = CopyStableSource(
                source,
                target,
                effectiveMaximum,
                limitMessage,
                cancellationToken,
                out copiedBytes);
            if (usesPhysicalIdentity && !OfficePathIdentity.IsOpenedFileWithinRootByIdentity(
                    sourcePath,
                    sourceDirectoryIdentity,
                    source.SafeFileHandle)) {
                throw new InvalidDataException(
                    "An external provenance manifest changed identity while it was being captured.");
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

    private static FileStream OpenSnapshotSource(
        string path,
        string sourceDirectoryIdentity,
        bool usesPhysicalIdentity) => OpenRegularFileForRead(
            path,
            sourceDirectoryIdentity,
            usesPhysicalIdentity,
            "The provenance snapshot source could not be opened as a regular file.");

    private static FileStream OpenRegularFileForRead(
        string path,
        string sourceDirectoryIdentity,
        bool usesPhysicalIdentity,
        string failureMessage) {
        try {
            if (usesPhysicalIdentity) {
                return OfficePathIdentity.OpenRegularFileForRead(path, sourceDirectoryIdentity, 81920);
            }

            string fullPath = Path.GetFullPath(path);
            if (!IsWithinDirectory(fullPath, sourceDirectoryIdentity)) {
                throw new InvalidDataException("The opened filesystem entry resolves outside the source directory.");
            }
            FileAttributes attributes = File.GetAttributes(fullPath);
            if ((attributes & (FileAttributes.Directory | FileAttributes.ReparsePoint)) != 0) {
                throw new InvalidDataException("The filesystem entry is not a portable regular file.");
            }
            var stream = new FileStream(
                fullPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                81920,
                FileOptions.SequentialScan);
            if (!stream.CanSeek) {
                stream.Dispose();
                throw new InvalidDataException("The filesystem entry is not a seekable regular file.");
            }
            return stream;
        } catch (InvalidDataException) {
            throw;
        } catch (FileNotFoundException) {
            throw;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new InvalidDataException(failureMessage, exception);
        }
    }

    internal static byte[] CopyStableSource(
        Stream source,
        Stream destination,
        long maximumBytes,
        string limitMessage,
        CancellationToken cancellationToken,
        out long copiedBytes) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!source.CanRead || !source.CanSeek) {
            throw new ArgumentException("The snapshot source must be a readable, seekable stream.", nameof(source));
        }
        if (!destination.CanWrite) {
            throw new ArgumentException("The snapshot destination must be writable.", nameof(destination));
        }
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        if (string.IsNullOrWhiteSpace(limitMessage)) throw new ArgumentException("A limit message is required.", nameof(limitMessage));

        source.Position = 0;
        if (source.Length > maximumBytes) throw new InvalidDataException(limitMessage);

        using IncrementalHash copiedHash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[81920];
        copiedBytes = 0;
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            copiedBytes += read;
            if (copiedBytes > maximumBytes) throw new InvalidDataException(limitMessage);
            copiedHash.AppendData(buffer, 0, read);
            destination.Write(buffer, 0, read);
        }
        cancellationToken.ThrowIfCancellationRequested();
        byte[] sha256 = copiedHash.GetHashAndReset();

        if (source.Length != copiedBytes) {
            throw new InvalidDataException("The snapshot source changed length while it was being captured.");
        }
        source.Position = 0;
        byte[] verificationHash = ComputeHashBounded(
            source,
            maximumBytes,
            limitMessage,
            cancellationToken,
            out long verifiedBytes);
        if (verifiedBytes != copiedBytes || source.Length != copiedBytes || !FixedTimeEquals(verificationHash, sha256)) {
            throw new InvalidDataException("The snapshot source changed content while it was being captured.");
        }
        return sha256;
    }

    private static byte[] ComputeHashBounded(
        Stream stream,
        long maximumBytes,
        string limitMessage,
        CancellationToken cancellationToken,
        out long hashedBytes) {
        using IncrementalHash algorithm = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[81920];
        hashedBytes = 0;
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            hashedBytes += read;
            if (hashedBytes > maximumBytes) throw new InvalidDataException(limitMessage);
            algorithm.AppendData(buffer, 0, read);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return algorithm.GetHashAndReset();
    }

    private static byte[] ComputeHash(Stream stream, CancellationToken cancellationToken) {
        using IncrementalHash algorithm = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[81920];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            algorithm.AppendData(buffer, 0, read);
        }
        cancellationToken.ThrowIfCancellationRequested();
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
