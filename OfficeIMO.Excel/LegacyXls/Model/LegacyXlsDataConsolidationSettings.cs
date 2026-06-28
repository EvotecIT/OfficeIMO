namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents worksheet-level BIFF DCon data-consolidation settings.
    /// </summary>
    public sealed class LegacyXlsDataConsolidationSettings {
        /// <summary>
        /// Creates worksheet data-consolidation settings parsed from a BIFF DCon record.
        /// </summary>
        public LegacyXlsDataConsolidationSettings(
            LegacyXlsDataConsolidationFunction function,
            bool usesTopLabels,
            bool usesLeftLabels,
            bool linksToSourceData,
            ushort rawFunction,
            ushort optionFlags) {
            Function = function;
            UsesTopLabels = usesTopLabels;
            UsesLeftLabels = usesLeftLabels;
            LinksToSourceData = linksToSourceData;
            RawFunction = rawFunction;
            OptionFlags = optionFlags;
        }

        /// <summary>Gets the decoded data-consolidation aggregation function.</summary>
        public LegacyXlsDataConsolidationFunction Function { get; }

        /// <summary>Gets whether the first source row supplies category labels.</summary>
        public bool UsesTopLabels { get; }

        /// <summary>Gets whether the first source column supplies category labels.</summary>
        public bool UsesLeftLabels { get; }

        /// <summary>Gets whether the consolidation links to source data.</summary>
        public bool LinksToSourceData { get; }

        /// <summary>Gets the raw BIFF aggregation function identifier.</summary>
        public ushort RawFunction { get; }

        /// <summary>Gets the raw BIFF data-consolidation option flags.</summary>
        public ushort OptionFlags { get; }
    }
}
