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

        // These checks run for every package part and relationship target, so they walk the path instead of
        // splitting it into segment strings.
        private static bool IsDirectoryOrEmptySegmentTarget(string path) {
            if (path.Length == 0) return true;
            int lastSeparator = -1;
            for (int index = 0; index < path.Length; index++) {
                if (path[index] != '/') continue;
                // Every segment after a separator must be non-empty.
                if (index == path.Length - 1 || path[index + 1] == '/') return true;
                lastSeparator = index;
            }
            return IsDotSegment(path, lastSeparator + 1, path.Length - lastSeparator - 1);
        }

        private static bool IsCanonicalPartPath(string path) {
            if (path.Length < 2 || path[0] != '/' || IsDirectoryOrEmptySegmentTarget(path)) return false;
            int segmentStart = 1;
            for (int index = 1; index <= path.Length; index++) {
                if (index < path.Length && path[index] != '/') {
                    char value = path[index];
                    if (value == '\\' || char.IsControl(value)) return false;
                    continue;
                }
                if (IsDotSegment(path, segmentStart, index - segmentStart)) return false;
                segmentStart = index + 1;
            }
            return true;
        }

        private static bool IsDotSegment(string path, int start, int length) =>
            length == 1 && path[start] == '.'
            || length == 2 && path[start] == '.' && path[start + 1] == '.';

        /// <summary>Decodes each segment once; encoded separators, dot segments and controls cannot identify another part.</summary>
        internal static string? DecodePartIdentity(string path) {
            if (!IsCanonicalPartPath(path)) return null;
            // A canonical path without escapes is its own identity.
            if (path.IndexOf('%') < 0) return path;
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
