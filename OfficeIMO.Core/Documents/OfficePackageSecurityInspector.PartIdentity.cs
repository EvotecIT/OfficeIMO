using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO {
    public static partial class OfficePackageSecurityInspector {
        /// <summary>Retains physical names and unique once-decoded aliases without merging source entries.</summary>
        private sealed class ZipPartIdentityIndex {
            private readonly Dictionary<string, string?> _identities = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

            internal void Add(string physicalName) {
                if (!OfficeOpcPartNames.IsCanonicalPartPath(physicalName)) return;
                AddAlias(physicalName, physicalName);
                string? decoded = OfficeOpcPartNames.DecodePartIdentity(physicalName);
                if (decoded != null) AddAlias(decoded, physicalName);
            }

            internal string? Resolve(string reference) {
                string? decoded = OfficeOpcPartNames.DecodePartIdentity(reference);
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
    }
}
