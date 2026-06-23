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
            IReadOnlyList<string> entries) {
            Kind = kind;
            Entries = entries?.ToArray() ?? throw new ArgumentNullException(nameof(entries));
        }

        /// <summary>Gets the compound feature kind.</summary>
        public LegacyXlsCompoundFeatureRecordKind Kind { get; }

        /// <summary>Gets matching compound directory entry paths or names.</summary>
        public IReadOnlyList<string> Entries { get; }
    }
}
