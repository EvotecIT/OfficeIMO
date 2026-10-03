using System.IO.Compression;
using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.IWork.Internal;

/// <summary>Copies a bounded local bundle into a deterministic private transport archive.</summary>
/// <remarks>This copies captured files; it does not generate or edit native iWork document semantics.</remarks>
internal sealed class IWorkDirectoryPackageSnapshot {
    private readonly string _path;
    private readonly string _identity;
    private readonly IWorkReadOptions _options;
    private readonly long _maximumTransportBytes;
    private string[]? _members;

    internal IWorkDirectoryPackageSnapshot(string path, IWorkReadOptions options, long maximumTransportBytes) {
        _path = System.IO.Path.GetFullPath(path).TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar);
        _options = options.Clone();
        _maximumTransportBytes = maximumTransportBytes;
        if (maximumTransportBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumTransportBytes));
        using var handle = OfficePathIdentity.OpenDirectoryForIdentity(_path, out _);
        _identity = OfficePathIdentity.GetPhysicalIdentityKey(_path, handle);
        OfficePathIdentity.EnsurePathMatchesOpenedDirectory(_path, handle);
    }

    internal Stream OpenStream(CancellationToken token) {
        token.ThrowIfCancellationRequested();
        VerifyIdentity();
        IWorkPackageData package = IWorkContainerReader.ReadDirectorySnapshot(_path, _options, token);
        VerifyIdentity();
        string[] members = package.SourceFilePaths.OrderBy(path => path, StringComparer.Ordinal).ToArray();
        if (_members is not null && !_members.SequenceEqual(members, StringComparer.Ordinal))
            throw new IOException("The selected iWork directory package source membership changed during execution.");
        _members ??= members;
        var output = new OfficeBoundedMemoryStream(_maximumTransportBytes);
        try {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                foreach (IWorkPackageEntry entry in package.Entries) {
                    token.ThrowIfCancellationRequested();
                    ZipArchiveEntry destination = archive.CreateEntry(entry.Path, CompressionLevel.NoCompression);
                    destination.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
                    using Stream target = destination.Open();
                    for (int offset = 0; offset < entry.Length;) {
                        token.ThrowIfCancellationRequested();
                        int count = Math.Min(64 * 1024, entry.Length - offset);
                        target.Write(entry.Bytes, offset, count);
                        offset += count;
                    }
                }
            }
            token.ThrowIfCancellationRequested();
            VerifyIdentity();
            output.Position = 0;
            return output;
        } catch {
            output.Dispose();
            throw;
        }
    }

    internal bool CanPublish(string destination, bool isDirectory, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        VerifyIdentity();
        string? output = OfficeStorageIdentity.GetLocalPath(destination);
        if (output is null) return true;
        if (OfficePathIdentity.IsSameOrDescendant(output, _path) ||
            isDirectory && OfficePathIdentity.IsSameOrDescendant(_path, output)) return false;
        if (!isDirectory && File.Exists(output)) {
            foreach (string member in _members ?? throw new InvalidOperationException("The package has not been captured.")) {
                token.ThrowIfCancellationRequested();
                if (OfficePathIdentity.AreEquivalent(member, output)) return false;
            }
        }
        VerifyIdentity();
        return true;
    }

    private void VerifyIdentity() {
        if (OfficePathIdentity.GetPhysicalIdentityKey(_path) != _identity)
            throw new IOException("The selected iWork directory package was replaced during execution.");
    }
}
