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
    }
}
