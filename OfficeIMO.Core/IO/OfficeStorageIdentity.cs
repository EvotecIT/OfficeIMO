using System;
using System.IO;

namespace OfficeIMO.Internal {
    /// <summary>Separates local filesystem identities from opaque storage-provider locations.</summary>
    internal static class OfficeStorageIdentity {
        /// <summary>Normalizes a file path or absolute provider URI without converting a provider URI to a path.</summary>
        internal static string Normalize(string location) {
            if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException("A storage location is required.", nameof(location));
            if (TryGetProviderUri(location, out Uri? provider)) return provider!.AbsoluteUri;
            if (Uri.TryCreate(location, UriKind.Absolute, out Uri? uri) && uri.IsFile) return Path.GetFullPath(uri.LocalPath);
            return Path.GetFullPath(location);
        }

        /// <summary>Returns a local path only for a filesystem location.</summary>
        internal static string? GetLocalPath(string location) =>
            TryGetProviderUri(location, out _) ? null : Normalize(location);

        /// <summary>Produces a stable persistence key; provider paths remain case-sensitive on every OS.</summary>
        internal static string GetPersistenceKey(string location) {
            string normalized = Normalize(location);
            return !TryGetProviderUri(normalized, out _) &&
                System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows)
                ? normalized.ToUpperInvariant() : normalized;
        }

        /// <summary>Compares provider locations exactly and delegates local aliases to the filesystem identity owner.</summary>
        internal static bool AreEquivalent(string first, string second) {
            string? firstPath = GetLocalPath(first);
            string? secondPath = GetLocalPath(second);
            if (firstPath is not null && secondPath is not null) return OfficePathIdentity.AreEquivalent(firstPath, secondPath);
            return firstPath is null && secondPath is null && string.Equals(Normalize(first), Normalize(second), StringComparison.Ordinal);
        }

        internal static string GetFileName(string location) {
            if (!TryGetProviderUri(location, out Uri? provider)) return Path.GetFileName(Normalize(location));
            string path = provider!.GetComponents(UriComponents.Path, UriFormat.Unescaped).TrimEnd('/');
            int separator = path.LastIndexOf('/');
            return path.Length > 0 ? path.Substring(separator + 1) : provider.Host;
        }

        private static bool TryGetProviderUri(string location, out Uri? uri) {
            // A drive-qualified Windows path is a file location, including when read on another host.
            if (location.Length >= 2 && char.IsLetter(location[0]) && location[1] == ':') {
                uri = null;
                return false;
            }
            return Uri.TryCreate(location, UriKind.Absolute, out uri) && !uri.IsFile;
        }
    }
}
