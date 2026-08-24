using System;

namespace OfficeIMO.Internal {
    /// <summary>Describes the physical identity and link metadata of one opened filesystem entry.</summary>
    internal readonly struct OfficeFileMetadata {
        internal OfficeFileMetadata(OfficePhysicalFileIdentity identity, ulong linkCount, uint unixMode,
            bool isDirectory) {
            Identity = identity;
            LinkCount = linkCount;
            UnixMode = unixMode;
            IsDirectory = isDirectory;
        }

        internal OfficePhysicalFileIdentity Identity { get; }
        internal ulong LinkCount { get; }
        internal uint UnixMode { get; }
        internal bool IsDirectory { get; }
    }

    /// <summary>Identifies one file within a filesystem authority and volume.</summary>
    internal readonly struct OfficePhysicalFileIdentity : IEquatable<OfficePhysicalFileIdentity> {
        internal OfficePhysicalFileIdentity(string authority, ulong volume, ulong fileLow, ulong fileHigh = 0) {
            Authority = authority ?? string.Empty;
            Volume = volume;
            FileLow = fileLow;
            FileHigh = fileHigh;
        }

        internal string Authority { get; }
        internal ulong Volume { get; }
        internal ulong FileLow { get; }
        internal ulong FileHigh { get; }

        internal bool HasSameNumericIdentity(OfficePhysicalFileIdentity other) =>
            Volume == other.Volume && FileLow == other.FileLow && FileHigh == other.FileHigh;

        public bool Equals(OfficePhysicalFileIdentity other) =>
            HasSameNumericIdentity(other) &&
            string.Equals(Authority, other.Authority, StringComparison.OrdinalIgnoreCase);

        public override bool Equals(object? obj) => obj is OfficePhysicalFileIdentity other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = StringComparer.OrdinalIgnoreCase.GetHashCode(Authority);
                hash = (hash * 397) ^ Volume.GetHashCode();
                hash = (hash * 397) ^ FileLow.GetHashCode();
                return (hash * 397) ^ FileHigh.GetHashCode();
            }
        }
    }
}
