namespace OfficeIMO.Excel {
    /// <summary>
    /// Configures package-native tabular XLSX exports.
    /// </summary>
    public sealed class ExcelWriterOptions {
        /// <summary>
        /// Rejects settings that require materializing the source rows before writing.
        /// For data readers, disable shared strings and automatic sizing.
        /// The workbook writer bounds row ingestion, including headerless tables. Explicit final save policies may require a resource-limited document finalization phase.
        /// </summary>
        public bool RequireStreaming { get; set; }

        /// <summary>Worksheet name.</summary>
        public string SheetName { get; set; } = "Data";

        /// <summary>Top-left cell for streamed output, including an optional title. Defaults to A1.</summary>
        public string Address { get; set; } = "A1";

        /// <summary>Optional bold title placed immediately above the exported headers or data.</summary>
        public string? Title { get; set; }

        /// <summary>Uses a bold font for exported headers.</summary>
        public bool BoldHeaders { get; set; }

        /// <summary>Freezes rows through the exported header, accounting for placement and title.</summary>
        public bool FreezeHeader { get; set; }

        /// <summary>Freezes columns through the first exported column.</summary>
        public bool FreezeFirstColumn { get; set; }

        /// <summary>Rejects an existing file before consuming rows and atomically refuses a destination created during export.</summary>
        public bool NoClobber { get; set; }

        /// <summary>Native header-based number formats resolved against the final table schema.</summary>
        public ExcelColumnFormatPlan? ColumnFormat { get; set; }

        /// <summary>Skips number-format rules whose headers are absent from the final schema.</summary>
        public bool IgnoreMissingColumnFormat { get; set; }

        /// <summary>Supplied core or application document properties, including the Author and Application aliases.</summary>
        public IReadOnlyDictionary<string, object?>? DocumentProperties { get; set; }

        /// <summary>Final native save policy. Formula evaluation, validation and encryption may require a resource-limited finalization phase after bounded row ingestion.</summary>
        public ExcelSaveOptions? SaveOptions { get; set; }

        /// <summary>Optional first-column emphasis on a streamed table.</summary>
        public bool ShowFirstColumn { get; set; }

        /// <summary>Optional last-column emphasis on a streamed table.</summary>
        public bool ShowLastColumn { get; set; }

        /// <summary>Enables alternating table row stripes. Defaults to true.</summary>
        public bool ShowRowStripes { get; set; } = true;

        /// <summary>Enables alternating table column stripes.</summary>
        public bool ShowColumnStripes { get; set; }

        /// <summary>Writes column names as the first worksheet row.</summary>
        public bool IncludeHeaders { get; set; } = true;

        /// <summary>Creates an Excel table over the exported range.</summary>
        public bool CreateTable { get; set; }

        /// <summary>Optional Excel table name.</summary>
        public string? TableName { get; set; }

        /// <summary>Excel table style used when <see cref="CreateTable"/> is enabled.</summary>
        public ExcelTableStyle TableStyle { get; set; } = ExcelTableStyle.TableStyleMedium2;

        /// <summary>Includes table filter buttons when a table is created.</summary>
        public bool IncludeAutoFilter { get; set; } = true;

        /// <summary>Calculates worksheet column widths from exported values.</summary>
        public bool AutoFit { get; set; }

        /// <summary>Uses the workbook's CellValue date and time number formats.</summary>
        public bool UseCellValueNumberFormats { get; set; }

        /// <summary>Writes explicit row and cell references. Disable for a smaller contiguous worksheet package.</summary>
        public bool IncludeCellReferences { get; set; } = true;

        /// <summary>
        /// Requests shared-string storage. The bounded writer rejects this setting; use an editable
        /// workbook for shared-string planning. Defaults to false (inline text).
        /// </summary>
        public bool UseSharedStrings { get; set; }

        /// <summary>Excel date system used for temporal values.</summary>
        public ExcelDateSystem DateSystem { get; set; } = ExcelDateSystem.NineteenHundred;

        /// <summary>
        /// Named complete styles, copied and compiled before reading source rows or preparing the
        /// first row. Styles cannot be registered during a row callback. Memory depends on the
        /// declared catalog and columns.
        /// </summary>
        public IReadOnlyDictionary<string, ExcelStyleDefinition>? Styles { get; set; }

        /// <summary>
        /// Optional declared style used as the default for each data row. Headers are excluded.
        /// Row defaults override column defaults as complete styles, without merging properties.
        /// </summary>
        public string? DefaultRowStyle { get; set; }

        /// <summary>
        /// Declared defaults for 1-based columns within the exported schema. Column formatting
        /// applies to the entire worksheet column, including headers and cells outside the data range.
        /// </summary>
        public IReadOnlyDictionary<int, string>? ColumnStyles { get; set; }

        internal ExcelTabularStylePlan? StylePlan { get; set; }
    }
}
