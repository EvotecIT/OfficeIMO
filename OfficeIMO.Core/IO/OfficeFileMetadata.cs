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
        internal bool IsRegularFile => UnixMode == 0 ? !IsDirectory : (UnixMode & 0xf000U) == 0x8000U;
    }

    /// <summary>Identifies one file within a filesystem authority and volume.</summary>
    internal readonly struct OfficePhysicalFileIdentity : IEquatable<OfficePhysicalFileIdentity> {
        private enum IdentityKind : byte {
            Native,
            WindowsExtended,
            WindowsLegacy
        }

        internal OfficePhysicalFileIdentity(string authority, ulong volume, ulong fileLow, ulong fileHigh = 0) {
            Authority = authority ?? string.Empty;
            Volume = volume;
            FileLow = fileLow;
            FileHigh = fileHigh;
            LegacyFileLow = 0;
            Kind = IdentityKind.Native;
        }

        private OfficePhysicalFileIdentity(string authority, ulong volume, ulong fileLow, ulong fileHigh,
            ulong legacyFileLow, IdentityKind kind) {
            Authority = authority ?? string.Empty;
            Volume = volume;
            FileLow = fileLow;
            FileHigh = fileHigh;
            LegacyFileLow = legacyFileLow;
            Kind = kind;
        }

        internal static OfficePhysicalFileIdentity CreateWindowsExtended(string authority, ulong volume,
            ulong fileLow, ulong fileHigh, ulong legacyFileLow) =>
            new OfficePhysicalFileIdentity(authority, volume, fileLow, fileHigh, legacyFileLow,
                IdentityKind.WindowsExtended);

        internal static OfficePhysicalFileIdentity CreateWindowsLegacy(string authority, ulong volume,
            ulong legacyFileLow) =>
            new OfficePhysicalFileIdentity(authority, volume, legacyFileLow, 0, legacyFileLow,
                IdentityKind.WindowsLegacy);

        internal string Authority { get; }
        internal ulong Volume { get; }
        internal ulong FileLow { get; }
        internal ulong FileHigh { get; }
        private ulong LegacyFileLow { get; }
        private IdentityKind Kind { get; }

        internal bool HasSameNumericIdentity(OfficePhysicalFileIdentity other) {
            if (Volume != other.Volume) return false;
            bool windows = Kind != IdentityKind.Native;
            bool otherWindows = other.Kind != IdentityKind.Native;
            if (windows != otherWindows) return false;
            if (windows) {
                if (Kind == IdentityKind.WindowsExtended && other.Kind == IdentityKind.WindowsExtended) {
                    return FileLow == other.FileLow && FileHigh == other.FileHigh;
                }
                return LegacyFileLow == other.LegacyFileLow;
            }
            return FileLow == other.FileLow && FileHigh == other.FileHigh;
        }

        internal bool HasSameAuthority(OfficePhysicalFileIdentity other) =>
            string.Equals(Authority, other.Authority, StringComparison.OrdinalIgnoreCase);

        internal string ToStableKey() {
            if (Kind == IdentityKind.WindowsExtended) {
                return "W128|" + Authority.ToUpperInvariant() + "|" + Volume.ToString("X16") + "|" +
                    FileHigh.ToString("X16") + "|" + FileLow.ToString("X16");
            }
            if (Kind == IdentityKind.WindowsLegacy) {
                return "W64|" + Authority.ToUpperInvariant() + "|" + Volume.ToString("X16") + "|" +
                    LegacyFileLow.ToString("X16");
            }
            return "N|" + Authority.ToUpperInvariant() + "|" + Volume.ToString("X16") + "|" +
                FileHigh.ToString("X16") + "|" + FileLow.ToString("X16");
        }

        public bool Equals(OfficePhysicalFileIdentity other) {
            if (Kind != other.Kind || Volume != other.Volume || !HasSameAuthority(other)) return false;
            if (Kind == IdentityKind.WindowsLegacy) return LegacyFileLow == other.LegacyFileLow;
            return FileLow == other.FileLow && FileHigh == other.FileHigh;
        }

        public override bool Equals(object? obj) => obj is OfficePhysicalFileIdentity other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = StringComparer.OrdinalIgnoreCase.GetHashCode(Authority);
                hash = (hash * 397) ^ Volume.GetHashCode();
                hash = (hash * 397) ^ Kind.GetHashCode();
                if (Kind == IdentityKind.WindowsLegacy) return (hash * 397) ^ LegacyFileLow.GetHashCode();
                hash = (hash * 397) ^ FileLow.GetHashCode();
                return (hash * 397) ^ FileHigh.GetHashCode();
            }
        }
    }
}
