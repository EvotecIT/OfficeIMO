namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a preserve-only feature discovered in the XLS compound container.
    /// </summary>
    public sealed class LegacyXlsCompoundFeatureRecord {
        /// <summary>
        /// Creates compound feature metadata.
        /// </summary>
        public LegacyXlsCompoundFeatureRecord(
            LegacyXlsCompoundFeatureRecordKind kind,
            IReadOnlyList<string> entries,
            IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole>? entryRoles = null) {
            Kind = kind;
            Entries = entries?.ToArray() ?? throw new ArgumentNullException(nameof(entries));
            var roles = new Dictionary<string, LegacyXlsCompoundFeatureEntryRole>(StringComparer.OrdinalIgnoreCase);
            if (entryRoles != null) {
                foreach (KeyValuePair<string, LegacyXlsCompoundFeatureEntryRole> role in entryRoles) {
                    roles[role.Key] = role.Value;
                }
            }

            EntryRoles = roles;
        }

        /// <summary>Gets the compound feature kind.</summary>
        public LegacyXlsCompoundFeatureRecordKind Kind { get; }

        /// <summary>Gets matching compound directory entry paths or names.</summary>
        public IReadOnlyList<string> Entries { get; }

        /// <summary>Gets matching compound directory entry roles keyed by entry path or name.</summary>
        public IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole> EntryRoles { get; }

        /// <summary>Gets VBA module stream names discovered in this compound feature.</summary>
        public IReadOnlyList<string> VbaModuleNames => Entries
            .Where(entry => EntryRoles.TryGetValue(entry, out LegacyXlsCompoundFeatureEntryRole role)
                && role == LegacyXlsCompoundFeatureEntryRole.VbaModuleStream)
            .Select(GetEntryLeafName)
            .ToArray();

        /// <summary>Gets the number of VBA module streams discovered in this compound feature.</summary>
        public int VbaModuleCount => VbaModuleNames.Count;

        private static string GetEntryLeafName(string entry) {
            int slashIndex = entry.LastIndexOf('/');
            int backslashIndex = entry.LastIndexOf('\\');
            int separatorIndex = Math.Max(slashIndex, backslashIndex);
            return separatorIndex >= 0 && separatorIndex + 1 < entry.Length ? entry.Substring(separatorIndex + 1) : entry;
        }
    }
}
