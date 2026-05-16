namespace OfficeIMO.Excel {
    /// <summary>
    /// Options controlling worksheet protection behavior.
    /// </summary>
    public sealed class ExcelSheetProtectionOptions {
        /// <summary>Allow selecting locked cells.</summary>
        public bool AllowSelectLockedCells { get; set; } = true;
        /// <summary>Allow selecting unlocked cells.</summary>
        public bool AllowSelectUnlockedCells { get; set; } = true;
        /// <summary>Allow formatting cells.</summary>
        public bool AllowFormatCells { get; set; }
        /// <summary>Allow formatting columns.</summary>
        public bool AllowFormatColumns { get; set; }
        /// <summary>Allow formatting rows.</summary>
        public bool AllowFormatRows { get; set; }
        /// <summary>Allow inserting columns.</summary>
        public bool AllowInsertColumns { get; set; }
        /// <summary>Allow inserting rows.</summary>
        public bool AllowInsertRows { get; set; }
        /// <summary>Allow inserting hyperlinks.</summary>
        public bool AllowInsertHyperlinks { get; set; }
        /// <summary>Allow deleting columns.</summary>
        public bool AllowDeleteColumns { get; set; }
        /// <summary>Allow deleting rows.</summary>
        public bool AllowDeleteRows { get; set; }
        /// <summary>Allow sorting.</summary>
        public bool AllowSort { get; set; }
        /// <summary>Allow AutoFilter.</summary>
        public bool AllowAutoFilter { get; set; }
        /// <summary>Allow PivotTables.</summary>
        public bool AllowPivotTables { get; set; }
        /// <summary>Optional worksheet protection password. This is Excel UI protection, not package encryption.</summary>
        public string? Password { get; set; }
        /// <summary>Optional precomputed legacy worksheet protection hash. When set, this value is written as-is.</summary>
        public string? LegacyPasswordHash { get; set; }
    }
}
