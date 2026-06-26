namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the OLE compound directory object type behind a preserve-only XLS feature entry.
    /// </summary>
    public enum LegacyXlsCompoundFeatureEntryObjectType {
        /// <summary>The entry object type is not known.</summary>
        Unknown = 0,

        /// <summary>The entry is an OLE compound storage.</summary>
        Storage = 1,

        /// <summary>The entry is an OLE compound stream.</summary>
        Stream = 2,

        /// <summary>The entry is the root OLE compound storage.</summary>
        RootStorage = 5
    }
}
