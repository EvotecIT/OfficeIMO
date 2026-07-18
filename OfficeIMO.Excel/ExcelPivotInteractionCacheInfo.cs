namespace OfficeIMO.Excel {
    /// <summary>
    /// Kind of workbook-level pivot interaction cache metadata.
    /// </summary>
    public enum ExcelPivotInteractionCacheKind {
        /// <summary>Slicer cache metadata.</summary>
        Slicer,
        /// <summary>Timeline cache metadata.</summary>
        Timeline
    }

    /// <summary>
    /// Describes workbook-level slicer or timeline cache metadata.
    /// </summary>
    public sealed class ExcelPivotInteractionCacheInfo {
        internal ExcelPivotInteractionCacheInfo(
            ExcelPivotInteractionCacheKind kind,
            string name,
            string? sourceName,
            string? pivotTableName,
            string relationshipId) {
            Kind = kind;
            Name = name;
            SourceName = sourceName;
            PivotTableName = pivotTableName;
            RelationshipId = relationshipId;
        }

        /// <summary>Interaction cache kind.</summary>
        public ExcelPivotInteractionCacheKind Kind { get; }

        /// <summary>Cache name.</summary>
        public string Name { get; }

        /// <summary>Source pivot field or table column name.</summary>
        public string? SourceName { get; }

        /// <summary>Bound pivot table name when declared.</summary>
        public string? PivotTableName { get; }

        /// <summary>Workbook relationship identifier for the cache part.</summary>
        public string RelationshipId { get; }
    }
}
