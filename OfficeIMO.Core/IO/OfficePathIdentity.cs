using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeIMO.Internal {
    /// <summary>Provides strict physical, link-aware path identity for trusted in-process consumers.</summary>
    internal static partial class OfficePathIdentity {
        internal static string Normalize(string path) {
            string identity = ResolvePhysicalPath(path);
            return Normalize(identity, IsCaseInsensitiveFileSystem(identity));
        }

        internal static string ResolvePhysicalPath(string path) {
            string fullPath = ResolveLinkSegments(Path.GetFullPath(path));
            string existingPath = TrimEndingDirectorySeparators(fullPath);
            var missingSegments = new Stack<string>();
            while (!TryGetPathMetadata(existingPath, out _)) {
                if (HasUnresolvedLinkEntry(existingPath)) {
                    throw new IOException("Could not safely resolve linked path '" + existingPath + "'.");
                }
                string? parent = Path.GetDirectoryName(existingPath);
                if (string.IsNullOrEmpty(parent) || string.Equals(parent, existingPath, StringComparison.Ordinal)) {
                    return fullPath;
                }
                string name = Path.GetFileName(existingPath);
                if (!string.IsNullOrEmpty(name)) missingSegments.Push(name);
                existingPath = parent;
            }

            string resolvedPath = ResolveExistingPhysicalPath(existingPath);
            foreach (string segment in missingSegments) resolvedPath = Path.Combine(resolvedPath, segment);
            return TrimEndingDirectorySeparators(Path.GetFullPath(resolvedPath));
        }

        internal static bool AreEquivalent(string left, string right) {
            string leftPath = ResolvePhysicalPath(left);
            string rightPath = ResolvePhysicalPath(right);
            if (TryGetPathAnchor(leftPath, out OfficePhysicalFileIdentity leftAnchor,
                    out string leftTail, out string leftExisting) &&
                TryGetPathAnchor(rightPath, out OfficePhysicalFileIdentity rightAnchor,
                    out string rightTail, out string rightExisting)) {
                if (!AreIdentitiesEquivalent(leftAnchor, rightAnchor)) return false;
                bool caseInsensitive = IsPotentiallyCaseInsensitiveFileSystem(leftExisting) ||
                    IsPotentiallyCaseInsensitiveFileSystem(rightExisting);
                return string.Equals(NormalizeMissingTail(leftTail, caseInsensitive),
                    NormalizeMissingTail(rightTail, caseInsensitive), StringComparison.Ordinal);
            }

            return string.Equals(Normalize(leftPath), Normalize(rightPath), StringComparison.Ordinal);
        }

        internal static string Normalize(string path, bool caseInsensitive) {
            string identity = Path.GetFullPath(path);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) identity = identity.Normalize(NormalizationForm.FormC);
            return caseInsensitive ? identity.ToUpperInvariant() : identity;
        }

        internal static StringComparer GetComparer(string path) =>
            IsCaseInsensitiveFileSystem(path) ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;

        internal static StringComparison GetComparison(string path) =>
            IsCaseInsensitiveFileSystem(path) ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

        internal static bool IsSameOrDescendant(string candidatePath, string rootPath) {
            string candidate = ResolvePhysicalPath(candidatePath);
            string root = ResolvePhysicalPath(rootPath);
            if (TryGetPathMetadata(root, out OfficeFileMetadata rootMetadata) &&
                ExistingAncestryContainsIdentity(candidate, rootMetadata.Identity)) {
                return true;
            }

            bool caseInsensitive = IsPotentiallyCaseInsensitiveFileSystem(root);
            string normalizedCandidate = Normalize(candidate, caseInsensitive);
            string normalizedRoot = TrimEndingDirectorySeparators(Normalize(root, caseInsensitive));
            if (string.Equals(normalizedCandidate, normalizedRoot, StringComparison.Ordinal)) return true;
            return normalizedCandidate.StartsWith(normalizedRoot + Path.DirectorySeparatorChar,
                StringComparison.Ordinal);
        }

        internal static bool HasMultipleLinks(string path) {
            return TryGetPathMetadata(Path.GetFullPath(path), out OfficeFileMetadata metadata) &&
                metadata.LinkCount > 1;
        }

        internal static string GetPhysicalIdentityKey(string path) =>
            GetMetadata(ResolvePhysicalPath(path)).Identity.ToStableKey();

        internal static string GetPhysicalIdentityKey(string path, SafeFileHandle handle) =>
            GetMetadata(path, handle).Identity.ToStableKey();

        internal static OfficeFileMetadata GetMetadata(string path) {
            string fullPath = Path.GetFullPath(path);
            if (!TryGetPathMetadata(fullPath, out OfficeFileMetadata metadata)) {
                throw new FileNotFoundException("The filesystem entry does not exist.", fullPath);
            }
            return metadata;
        }

        internal static OfficeFileMetadata GetMetadata(string path, SafeFileHandle handle) {
            if (handle == null || handle.IsInvalid || handle.IsClosed) {
                throw new IOException("The filesystem handle is not available for identity inspection.");
            }
            string fullPath = Path.GetFullPath(path);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return GetWindowsMetadata(fullPath, handle);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) || RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                return GetUnixMetadata(handle);
            }
            throw new PlatformNotSupportedException("Physical file identity is not supported on this platform.");
        }

        internal static FileStream OpenRegularFileForRead(string path, string physicalRoot, int bufferSize) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            if (physicalRoot == null) throw new ArgumentNullException(nameof(physicalRoot));
            if (bufferSize <= 0) throw new ArgumentOutOfRangeException(nameof(bufferSize));
            string fullPath = Path.GetFullPath(path);
            FileStream stream;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                stream = OpenWindowsRegularFileForRead(fullPath, bufferSize);
            } else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) || RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                stream = OpenUnixRegularFileForRead(fullPath, bufferSize);
            } else {
                throw new PlatformNotSupportedException("Regular-file opening is not supported on this platform.");
            }
            try {
                string? openedPath = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                    ? GetWindowsFinalPath(stream.SafeFileHandle)
                    : GetUnixOpenedPath(stream.SafeFileHandle);
                bool isWithinRoot = openedPath != null
                    ? IsPhysicalPathWithinRoot(openedPath, physicalRoot)
                    : RuntimeInformation.IsOSPlatform(OSPlatform.Linux)
                        && IsOpenedFileWithinRootByIdentity(fullPath, physicalRoot, stream.SafeFileHandle);
                if (!isWithinRoot) {
                    throw new InvalidDataException("The opened filesystem entry resolves outside the source directory.");
                }
                return stream;
            } catch {
                stream.Dispose();
                throw;
            }
        }

        internal static bool IsOpenedFileWithinRootByIdentity(string path, string physicalRoot,
            SafeFileHandle handle) {
            string resolvedPath = ResolvePhysicalPath(path);
            if (!IsPhysicalPathWithinRoot(resolvedPath, physicalRoot)) return false;
            OfficeFileMetadata pathMetadata = GetMetadata(resolvedPath);
            OfficeFileMetadata handleMetadata = GetMetadata(resolvedPath, handle);
            return AreIdentitiesEquivalent(pathMetadata.Identity, handleMetadata.Identity);
        }

        private static bool IsPhysicalPathWithinRoot(string candidatePath, string physicalRoot) {
            string root = TrimEndingDirectorySeparators(Path.GetFullPath(physicalRoot));
            bool caseInsensitive = IsCaseInsensitiveFileSystem(root);
            string candidate = Normalize(candidatePath, caseInsensitive);
            string normalizedRoot = TrimEndingDirectorySeparators(Normalize(root, caseInsensitive));
            return candidate.StartsWith(normalizedRoot + Path.DirectorySeparatorChar, StringComparison.Ordinal);
        }

        internal static bool IsCaseInsensitiveFileSystem(string path) {
            return TryGetFileSystemCaseBehavior(path, out bool caseInsensitive)
                ? caseInsensitive
                : IsConservativelyCaseInsensitivePlatform;
        }

        private static bool IsPotentiallyCaseInsensitiveFileSystem(string path) =>
            !TryGetFileSystemCaseBehavior(path, out bool caseInsensitive) || caseInsensitive;

        private static bool TryGetFileSystemCaseBehavior(string path, out bool caseInsensitive) {
            caseInsensitive = false;
            string fullPath = Path.GetFullPath(path);
            string? existingPath = TrimEndingDirectorySeparators(fullPath);
            OfficeFileMetadata existingMetadata = default(OfficeFileMetadata);
            while (!string.IsNullOrEmpty(existingPath) &&
                   !TryGetPathMetadata(existingPath, out existingMetadata)) {
                existingPath = Path.GetDirectoryName(existingPath);
            }
            if (string.IsNullOrEmpty(existingPath)) return false;

            string? directory = existingMetadata.IsDirectory ? existingPath : Path.GetDirectoryName(existingPath);
            if (string.IsNullOrEmpty(directory)) return false;

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) &&
                TryGetWindowsDirectoryCaseInsensitive(directory, out caseInsensitive)) return true;
            if (TryGetPathMetadata(directory, out OfficeFileMetadata directoryMetadata) && directoryMetadata.IsDirectory) {
                if ((RuntimeInformation.IsOSPlatform(OSPlatform.Linux) ||
                     RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) &&
                    TryGetUnixDirectoryCaseInsensitive(directory, out caseInsensitive)) return true;
                try {
                    foreach (string entry in Directory.EnumerateFileSystemEntries(directory))
                        if (TryDetectFromExistingName(entry, out caseInsensitive)) return true;
                } catch (IOException) {
                } catch (UnauthorizedAccessException) {
                }
            }
            return false;
        }

        private static bool IsConservativelyCaseInsensitivePlatform =>
            RuntimeInformation.IsOSPlatform(OSPlatform.Windows) || RuntimeInformation.IsOSPlatform(OSPlatform.OSX);

        private static string ResolveExistingPhysicalPath(string path) {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return ResolveWindowsExistingPath(path);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) || RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                return ResolveUnixExistingPath(path);
            }
            throw new PlatformNotSupportedException("Physical path resolution is not supported on this platform.");
        }

        private static bool TryDetectFromExistingName(string path, out bool caseInsensitive) {
            caseInsensitive = false;
            string existingPath = TrimEndingDirectorySeparators(path);
            string? parent = Path.GetDirectoryName(existingPath);
            string name = Path.GetFileName(existingPath);
            if (string.IsNullOrEmpty(parent) || string.IsNullOrEmpty(name)) return false;

            int letterIndex = -1;
            for (int index = name.Length - 1; index >= 0; index--) {
                if (char.IsLetter(name[index])) { letterIndex = index; break; }
            }
            if (letterIndex < 0) return false;

            char original = name[letterIndex];
            char alternate = char.IsUpper(original) ? char.ToLowerInvariant(original) : char.ToUpperInvariant(original);
            if (alternate == original) return false;
            var alternateName = new StringBuilder(name);
            alternateName[letterIndex] = alternate;
            string alternatePath = Path.Combine(parent, alternateName.ToString());
            if (!TryGetPathMetadata(alternatePath, out _)) return true;

            try {
                int matches = Directory.EnumerateFileSystemEntries(parent)
                    .Count(entry => string.Equals(Path.GetFileName(entry), name, StringComparison.OrdinalIgnoreCase));
                caseInsensitive = matches <= 1;
                return true;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            }
        }

        private static bool ExistingAncestryContainsIdentity(string path, OfficePhysicalFileIdentity expected) {
            string? candidate = TrimEndingDirectorySeparators(Path.GetFullPath(path));
            while (!string.IsNullOrEmpty(candidate)) {
                if (TryGetPathMetadata(candidate, out OfficeFileMetadata observed) &&
                    AreIdentitiesEquivalent(observed.Identity, expected)) return true;
                string? parent = Path.GetDirectoryName(candidate);
                if (string.IsNullOrEmpty(parent) || string.Equals(parent, candidate, StringComparison.Ordinal)) break;
                candidate = parent;
            }
            return false;
        }

        private static bool TryGetPathMetadata(string path, out OfficeFileMetadata metadata) {
            string fullPath = Path.GetFullPath(path);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                return TryGetWindowsMetadata(fullPath, out metadata);
            }
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) || RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                return TryGetUnixMetadata(fullPath, out metadata);
            }
            throw new PlatformNotSupportedException("Physical file identity is not supported on this platform.");
        }

        private static bool TryGetPathAnchor(string path, out OfficePhysicalFileIdentity identity,
            out string relativeTail, out string existingPath) {
            existingPath = TrimEndingDirectorySeparators(Path.GetFullPath(path));
            var missing = new Stack<string>();
            while (!TryGetPathMetadata(existingPath, out _)) {
                if (HasUnresolvedLinkEntry(existingPath)) {
                    throw new IOException("Could not safely inspect linked path '" + existingPath + "'.");
                }
                string? parent = Path.GetDirectoryName(existingPath);
                if (string.IsNullOrEmpty(parent) || string.Equals(parent, existingPath, StringComparison.Ordinal)) {
                    identity = default(OfficePhysicalFileIdentity);
                    relativeTail = string.Empty;
                    return false;
                }
                string name = Path.GetFileName(existingPath);
                if (!string.IsNullOrEmpty(name)) missing.Push(name);
                existingPath = parent;
            }
            identity = GetMetadata(existingPath).Identity;
            relativeTail = string.Join("/", missing.ToArray());
            return true;
        }

        private static bool AreIdentitiesEquivalent(OfficePhysicalFileIdentity left,
            OfficePhysicalFileIdentity right) {
            if (!left.HasSameNumericIdentity(right)) return false;
            if (left.HasSameAuthority(right)) return true;
            throw new IOException("Matching filesystem identifiers came from different authorities and cannot be compared safely.");
        }

        private static string NormalizeMissingTail(string tail, bool caseInsensitive) {
            string normalized = tail.Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/');
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) normalized = normalized.Normalize(NormalizationForm.FormC);
            return caseInsensitive ? normalized.ToUpperInvariant() : normalized;
        }

        private static string ResolveLinkSegments(string path) {
            string current = Path.GetFullPath(path);
            for (int pass = 0; pass < 40; pass++) {
                string resolved = ResolveLinkPass(current);
                if (string.Equals(current, resolved, StringComparison.Ordinal)) return resolved;
                current = resolved;
            }
            throw new IOException("Linked path resolution exceeded the supported depth.");
        }

        private static string ResolveLinkPass(string path) {
            string fullPath = Path.GetFullPath(path);
            string? root = Path.GetPathRoot(fullPath);
            if (string.IsNullOrEmpty(root)) return fullPath;
            string current = root!;
            string[] segments = fullPath.Substring(root!.Length).Split(
                new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
                StringSplitOptions.RemoveEmptyEntries);
            for (int index = 0; index < segments.Length; index++) {
                current = Path.Combine(current, segments[index]);
                if (TryReadLinkTarget(current, out string? target)) {
                    current = Path.GetFullPath(Path.IsPathRooted(target!)
                        ? target!
                        : Path.Combine(Path.GetDirectoryName(current)!, target!));
                    continue;
                }
                if (!TryGetPathMetadata(current, out _)) {
                    if (HasUnresolvedLinkEntry(current)) {
                        throw new IOException("Could not safely resolve linked path '" + current + "'.");
                    }
                    for (int remainder = index + 1; remainder < segments.Length; remainder++) {
                        current = Path.Combine(current, segments[remainder]);
                    }
                    break;
                }
            }
            return TrimEndingDirectorySeparators(Path.GetFullPath(current));
        }

        private static bool TryReadLinkTarget(string path, out string? target) {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return TryReadWindowsLinkTarget(path, out target);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) || RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                return TryReadUnixLinkTarget(path, out target);
            }
            target = null;
            return false;
        }

        private static bool HasUnresolvedLinkEntry(string path) {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return HasWindowsReparsePoint(path);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) || RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                return IsUnixLink(path);
            }
            return false;
        }

        private static string TrimEndingDirectorySeparators(string path) {
            string root = Path.GetPathRoot(path) ?? string.Empty;
            int length = path.Length;
            while (length > root.Length &&
                   (path[length - 1] == Path.DirectorySeparatorChar || path[length - 1] == Path.AltDirectorySeparatorChar)) {
                length--;
            }
            return length == path.Length ? path : path.Substring(0, length);
        }
    }
}
