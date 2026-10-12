namespace OfficeIMO.Excel {
    /// <summary>
    /// Configures package-native tabular XLSX exports.
    /// </summary>
    public sealed class ExcelTabularWriteOptions {
        /// <summary>
        /// Rejects settings that require materializing the source rows before writing.
        /// For data readers, disable shared strings and automatic sizing; tables require headers.
        /// For row sequences, disable tables and automatic sizing. Defaults to false.
        /// Shared-string row exports retain unique text while consuming each source row once.
        /// </summary>
        public bool RequireStreaming { get; set; }

        /// <summary>Worksheet name.</summary>
        public string SheetName { get; set; } = "Data";

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
        /// Stores text in the workbook shared-string table. Row writers honor this setting in a
        /// single source pass and retain unique text in memory. When row-write options are omitted,
        /// row writers use inline strings; data-reader exports enable shared-string planning by default.
        /// </summary>
        public bool UseSharedStrings { get; set; } = true;

        /// <summary>Excel date system used for temporal values.</summary>
        public ExcelDateSystem DateSystem { get; set; } = ExcelDateSystem.NineteenHundred;

        /// <summary>
        /// Named complete styles, copied and compiled before reading source rows or preparing the
        /// destination. Styles cannot be registered during a row callback. Memory depends on the
        /// declared catalog and columns, plus unique text when shared strings are enabled.
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
