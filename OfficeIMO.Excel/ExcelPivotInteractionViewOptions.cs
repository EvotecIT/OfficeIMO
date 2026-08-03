namespace OfficeIMO.Excel {
    /// <summary>Display and placement options for a native worksheet slicer.</summary>
    public sealed class ExcelSlicerViewOptions {
        /// <summary>Optional unique slicer view name. A stable name is generated when omitted.</summary>
        public string? Name { get; set; }

        /// <summary>Optional existing compatible cache name to reuse.</summary>
        public string? CacheName { get; set; }

        /// <summary>Optional displayed caption. The source field is used when omitted.</summary>
        public string? Caption { get; set; }

        /// <summary>Built-in slicer style.</summary>
        public string Style { get; set; } = "SlicerStyleLight2";

        /// <summary>One-based anchor row.</summary>
        public int Row { get; set; } = 1;

        /// <summary>One-based anchor column.</summary>
        public int Column { get; set; } = 1;

        /// <summary>Rendered width in pixels.</summary>
        public int WidthPixels { get; set; } = 180;

        /// <summary>Rendered height in pixels.</summary>
        public int HeightPixels { get; set; } = 240;

        /// <summary>Number of item columns displayed by the slicer.</summary>
        public int ItemColumns { get; set; } = 1;

        /// <summary>Whether the caption is displayed.</summary>
        public bool ShowCaption { get; set; } = true;

        /// <summary>Whether the slicer position is locked.</summary>
        public bool LockedPosition { get; set; }
    }

    /// <summary>Timeline display level.</summary>
    public enum ExcelTimelineLevel {
        /// <summary>Years.</summary>
        Year = 0,
        /// <summary>Quarters.</summary>
        Quarter = 1,
        /// <summary>Months.</summary>
        Month = 2,
        /// <summary>Days.</summary>
        Day = 3
    }

    /// <summary>Display and placement options for a native worksheet timeline.</summary>
    public sealed class ExcelTimelineViewOptions {
        /// <summary>Optional unique timeline view name. A stable name is generated when omitted.</summary>
        public string? Name { get; set; }

        /// <summary>Optional existing compatible cache name to reuse.</summary>
        public string? CacheName { get; set; }

        /// <summary>Optional displayed caption. The source field is used when omitted.</summary>
        public string? Caption { get; set; }

        /// <summary>Built-in timeline style.</summary>
        public string Style { get; set; } = "TimelineStyleLight2";

        /// <summary>One-based anchor row.</summary>
        public int Row { get; set; } = 1;

        /// <summary>One-based anchor column.</summary>
        public int Column { get; set; } = 1;

        /// <summary>Rendered width in pixels.</summary>
        public int WidthPixels { get; set; } = 480;

        /// <summary>Rendered height in pixels.</summary>
        public int HeightPixels { get; set; } = 120;

        /// <summary>Initial display level.</summary>
        public ExcelTimelineLevel Level { get; set; } = ExcelTimelineLevel.Month;

        /// <summary>Whether the header is displayed.</summary>
        public bool ShowHeader { get; set; } = true;

        /// <summary>Whether the selected range label is displayed.</summary>
        public bool ShowSelectionLabel { get; set; } = true;

        /// <summary>Whether the current time level is displayed.</summary>
        public bool ShowTimeLevel { get; set; } = true;

        /// <summary>Whether the horizontal scrollbar is displayed.</summary>
        public bool ShowHorizontalScrollbar { get; set; } = true;
    }

    /// <summary>Native slicer or timeline view bound to a PivotTable cache.</summary>
    public sealed class ExcelPivotInteractionInfo {
        internal ExcelPivotInteractionInfo(
            ExcelPivotInteractionCacheKind kind,
            string name,
            string cacheName,
            string sourceName,
            string? pivotTableName,
            string worksheetName,
            string relationshipId) {
            Kind = kind;
            Name = name;
            CacheName = cacheName;
            SourceName = sourceName;
            PivotTableName = pivotTableName;
            WorksheetName = worksheetName;
            RelationshipId = relationshipId;
        }

        /// <summary>Interaction kind.</summary>
        public ExcelPivotInteractionCacheKind Kind { get; }

        /// <summary>Unique view name.</summary>
        public string Name { get; }

        /// <summary>Native cache name.</summary>
        public string CacheName { get; }

        /// <summary>Pivot source field.</summary>
        public string SourceName { get; }

        /// <summary>Bound PivotTable name, when it can be resolved.</summary>
        public string? PivotTableName { get; }

        /// <summary>Worksheet that displays the view.</summary>
        public string WorksheetName { get; }

        /// <summary>Worksheet relationship identifier for the native view part.</summary>
        public string RelationshipId { get; }
    }
}
