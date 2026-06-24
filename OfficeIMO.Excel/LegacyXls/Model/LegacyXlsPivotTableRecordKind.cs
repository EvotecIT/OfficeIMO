namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies a preserve-only PivotTable BIFF record decoded into the legacy XLS import model.
    /// </summary>
    public enum LegacyXlsPivotTableRecordKind {
        /// <summary>PivotTable record that is currently preserved without record-specific field decoding.</summary>
        PreserveOnly,

        /// <summary>PivotTable view definition metadata.</summary>
        View,

        /// <summary>Pivot field metadata.</summary>
        Field,

        /// <summary>Pivot item metadata.</summary>
        Item,

        /// <summary>Pivot field index-list metadata.</summary>
        FieldIndexList,

        /// <summary>Pivot line item metadata.</summary>
        LineItem,

        /// <summary>Pivot page item metadata.</summary>
        PageItem,

        /// <summary>Data item metadata from an SXDI record.</summary>
        DataItem,

        /// <summary>Pivot cache or cache-source metadata.</summary>
        Cache,

        /// <summary>Pivot cache item value metadata.</summary>
        CacheItem,

        /// <summary>Pivot table layout metadata.</summary>
        Table,

        /// <summary>Pivot cache stream metadata.</summary>
        CacheStream,

        /// <summary>Pivot cache source-data metadata.</summary>
        CacheSource,

        /// <summary>Grouping range metadata from an SXRng record.</summary>
        GroupingRange,

        /// <summary>Pivot rule metadata.</summary>
        Rule,

        /// <summary>Pivot filter metadata.</summary>
        Filter,

        /// <summary>Pivot formatting metadata.</summary>
        Format,

        /// <summary>Pivot formula metadata.</summary>
        Formula,

        /// <summary>Pivot selection metadata.</summary>
        Selection,

        /// <summary>Extended pivot field metadata from an SXVDEx record.</summary>
        ExtendedPivotField,

        /// <summary>Pivot cache extension metadata.</summary>
        CacheExtension,

        /// <summary>Pivot query table tag metadata.</summary>
        QueryTableTag,

        /// <summary>Pivot view link metadata.</summary>
        ViewLink,

        /// <summary>Pivot chart linkage metadata.</summary>
        PivotChart,

        /// <summary>Additional pivot metadata extension record.</summary>
        Additional
    }
}
