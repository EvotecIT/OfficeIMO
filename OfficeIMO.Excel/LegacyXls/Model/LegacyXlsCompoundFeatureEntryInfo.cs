namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes one OLE compound directory entry that belongs to a preserve-only XLS feature.
    /// </summary>
    public sealed class LegacyXlsCompoundFeatureEntryInfo {
        /// <summary>
        /// Creates compound feature entry metadata.
        /// </summary>
        public LegacyXlsCompoundFeatureEntryInfo(
            string path,
            LegacyXlsCompoundFeatureEntryRole role,
            LegacyXlsCompoundFeatureEntryObjectType objectType = LegacyXlsCompoundFeatureEntryObjectType.Unknown,
            long? sizeBytes = null) {
            Path = string.IsNullOrWhiteSpace(path) ? throw new ArgumentException("Entry path must not be empty.", nameof(path)) : path;
            Role = role;
            ObjectType = objectType;
            SizeBytes = sizeBytes;
        }

        /// <summary>Gets the compound directory entry path or name.</summary>
        public string Path { get; }

        /// <summary>Gets the preserve-only role assigned to this compound entry.</summary>
        public LegacyXlsCompoundFeatureEntryRole Role { get; }

        /// <summary>Gets the OLE compound directory object type.</summary>
        public LegacyXlsCompoundFeatureEntryObjectType ObjectType { get; }

        /// <summary>Gets the stream or storage size declared by the OLE compound directory, when known.</summary>
        public long? SizeBytes { get; }

        /// <summary>Gets whether this entry is an OLE compound storage.</summary>
        public bool IsStorage => ObjectType == LegacyXlsCompoundFeatureEntryObjectType.Storage
            || ObjectType == LegacyXlsCompoundFeatureEntryObjectType.RootStorage;

        /// <summary>Gets whether this entry is an OLE compound stream.</summary>
        public bool IsStream => ObjectType == LegacyXlsCompoundFeatureEntryObjectType.Stream;
    }
}
