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
            IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole>? entryRoles = null,
            IReadOnlyDictionary<string, long>? entrySizes = null,
            IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryObjectType>? entryObjectTypes = null,
            IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryContentKind>? entryContentKinds = null) {
            Kind = kind;
            Entries = entries?.ToArray() ?? throw new ArgumentNullException(nameof(entries));
            var roles = new Dictionary<string, LegacyXlsCompoundFeatureEntryRole>(StringComparer.OrdinalIgnoreCase);
            if (entryRoles != null) {
                foreach (KeyValuePair<string, LegacyXlsCompoundFeatureEntryRole> role in entryRoles) {
                    roles[role.Key] = role.Value;
                }
            }

            EntryRoles = roles;
            var sizes = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            if (entrySizes != null) {
                foreach (KeyValuePair<string, long> size in entrySizes) {
                    sizes[size.Key] = size.Value;
                }
            }

            EntrySizes = sizes;
            var objectTypes = new Dictionary<string, LegacyXlsCompoundFeatureEntryObjectType>(StringComparer.OrdinalIgnoreCase);
            if (entryObjectTypes != null) {
                foreach (KeyValuePair<string, LegacyXlsCompoundFeatureEntryObjectType> objectType in entryObjectTypes) {
                    objectTypes[objectType.Key] = objectType.Value;
                }
            }

            EntryObjectTypes = objectTypes;
            var contentKinds = new Dictionary<string, LegacyXlsCompoundFeatureEntryContentKind>(StringComparer.OrdinalIgnoreCase);
            if (entryContentKinds != null) {
                foreach (KeyValuePair<string, LegacyXlsCompoundFeatureEntryContentKind> contentKind in entryContentKinds) {
                    contentKinds[contentKind.Key] = contentKind.Value;
                }
            }

            EntryContentKinds = contentKinds;
            EntryDetails = Entries
                .Select(entry => new LegacyXlsCompoundFeatureEntryInfo(
                    entry,
                    EntryRoles.TryGetValue(entry, out LegacyXlsCompoundFeatureEntryRole role) ? role : LegacyXlsCompoundFeatureEntryRole.Unknown,
                    EntryObjectTypes.TryGetValue(entry, out LegacyXlsCompoundFeatureEntryObjectType objectType) ? objectType : LegacyXlsCompoundFeatureEntryObjectType.Unknown,
                    EntrySizes.TryGetValue(entry, out long size) ? size : null,
                    EntryContentKinds.TryGetValue(entry, out LegacyXlsCompoundFeatureEntryContentKind contentKind) ? contentKind : LegacyXlsCompoundFeatureEntryContentKind.Unknown))
                .ToArray();
        }

        /// <summary>Gets the compound feature kind.</summary>
        public LegacyXlsCompoundFeatureRecordKind Kind { get; }

        /// <summary>Gets matching compound directory entry paths or names.</summary>
        public IReadOnlyList<string> Entries { get; }

        /// <summary>Gets matching compound directory entry roles keyed by entry path or name.</summary>
        public IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole> EntryRoles { get; }

        /// <summary>Gets matching compound directory entry declared sizes keyed by entry path or name.</summary>
        public IReadOnlyDictionary<string, long> EntrySizes { get; }

        /// <summary>Gets matching compound directory entry object types keyed by entry path or name.</summary>
        public IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryObjectType> EntryObjectTypes { get; }

        /// <summary>Gets matching compound directory entry content shapes keyed by entry path or name.</summary>
        public IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryContentKind> EntryContentKinds { get; }

        /// <summary>Gets typed metadata for matching compound directory entries.</summary>
        public IReadOnlyList<LegacyXlsCompoundFeatureEntryInfo> EntryDetails { get; }

        /// <summary>Gets the total declared byte size of matching compound directory entries with known sizes.</summary>
        public long EntryByteCount => EntryDetails
            .Where(entry => entry.SizeBytes.HasValue)
            .Sum(entry => entry.SizeBytes!.Value);

        /// <summary>Gets VBA module stream names discovered in this compound feature.</summary>
        public IReadOnlyList<string> VbaModuleNames => EntryDetails
            .Where(entry => entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaModuleStream)
            .Select(entry => GetEntryLeafName(entry.Path))
            .ToArray();

        /// <summary>Gets the number of VBA module streams discovered in this compound feature.</summary>
        public int VbaModuleCount => VbaModuleNames.Count;

        /// <summary>Gets the total declared byte size of VBA module streams with known sizes.</summary>
        public long VbaModuleByteCount => EntryDetails
            .Where(entry => entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaModuleStream && entry.SizeBytes.HasValue)
            .Sum(entry => entry.SizeBytes!.Value);

        private static string GetEntryLeafName(string entry) {
            int slashIndex = entry.LastIndexOf('/');
            int backslashIndex = entry.LastIndexOf('\\');
            int separatorIndex = Math.Max(slashIndex, backslashIndex);
            return separatorIndex >= 0 && separatorIndex + 1 < entry.Length ? entry.Substring(separatorIndex + 1) : entry;
        }
    }
}
