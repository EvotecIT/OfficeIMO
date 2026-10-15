using System;
using System.Collections.Generic;

namespace OfficeIMO {
    public static partial class OfficePackageSecurityInspector {
        /// <summary>Retains physical names and unique once-decoded aliases without merging source entries.</summary>
        private sealed class ZipPartIdentityIndex {
            private readonly Dictionary<string, string?> _identities = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

            internal void Add(string physicalName) {
                if (!IsCanonicalPartPath(physicalName)) return;
                AddAlias(physicalName, physicalName);
                string? decoded = DecodePartIdentity(physicalName);
                if (decoded != null) AddAlias(decoded, physicalName);
            }

            internal string? Resolve(string reference) {
                string? decoded = DecodePartIdentity(reference);
                if (decoded == null) return null;
                if (_identities.TryGetValue(decoded, out string? physicalName)) return physicalName;
                // Preserve a well-formed missing reference's existing identity;
                // a collision or malformed reference remains unresolved instead.
                return reference;
            }

            private void AddAlias(string alias, string physicalName) {
                if (!_identities.TryGetValue(alias, out string? existing)) {
                    _identities.Add(alias, physicalName);
                } else if (!string.Equals(existing, physicalName, StringComparison.Ordinal)) {
                    _identities[alias] = null;
                }
            }
        }

        private static bool IsDirectoryOrEmptySegmentTarget(string path) {
            if (path.Length == 0) return true;
            string[] segments = path.Split('/');
            for (int index = 1; index < segments.Length; index++) {
                if (segments[index].Length == 0) return true;
            }
            string last = segments[segments.Length - 1];
            return last == "." || last == "..";
        }

        private static bool IsCanonicalPartPath(string path) {
            if (path.Length < 2 || path[0] != '/' || IsDirectoryOrEmptySegmentTarget(path)) return false;
            foreach (string segment in path.Substring(1).Split('/')) {
                if (segment == "." || segment == ".." || segment.IndexOf('\\') >= 0) return false;
                foreach (char value in segment) if (char.IsControl(value)) return false;
            }
            return true;
        }

        /// <summary>Decodes each segment once; encoded separators, dot segments and controls cannot identify another part.</summary>
        private static string? DecodePartIdentity(string path) {
            if (!IsCanonicalPartPath(path)) return null;
            string[] segments = path.Split('/');
            for (int index = 1; index < segments.Length; index++) {
                string? decoded = DecodePartIdentitySegment(segments[index]);
                if (decoded == null) return null;
                segments[index] = decoded;
            }
            return string.Join("/", segments);
        }

        private static string? DecodePartIdentitySegment(string segment) {
            for (int offset = 0; offset < segment.Length; offset++) {
                if (segment[offset] != '%') continue;
                if (offset + 2 >= segment.Length || !IsPartIdentityHexDigit(segment[offset + 1])
                    || !IsPartIdentityHexDigit(segment[offset + 2])) return null;
                offset += 2;
            }
            string decoded;
            try { decoded = Uri.UnescapeDataString(segment); }
            catch (UriFormatException) { return null; }
            if (decoded.Length == 0 || decoded == "." || decoded == ".."
                || decoded.IndexOf('/') >= 0 || decoded.IndexOf('\\') >= 0) return null;
            foreach (char value in decoded) if (char.IsControl(value)) return null;
            return decoded;
        }

        private static bool IsPartIdentityHexDigit(char value) => value >= '0' && value <= '9'
            || value >= 'a' && value <= 'f' || value >= 'A' && value <= 'F';
    }
}
