namespace OfficeIMO.Excel {
    /// <summary>
    /// Configures package-native tabular XLSX exports.
    /// </summary>
    public sealed class ExcelTabularWriteOptions {
        /// <summary>Worksheet name.</summary>
        public string SheetName { get; set; } = "Data";

        /// <summary>Writes column names as the first worksheet row.</summary>
        public bool IncludeHeaders { get; set; } = true;

        /// <summary>Creates an Excel table over the exported range.</summary>
        public bool CreateTable { get; set; }

        /// <summary>Optional Excel table name.</summary>
        public string? TableName { get; set; }

        /// <summary>Excel table style used when <see cref="CreateTable"/> is enabled.</summary>
        public TableStyle TableStyle { get; set; } = TableStyle.TableStyleMedium2;

        /// <summary>Includes table filter buttons when a table is created.</summary>
        public bool IncludeAutoFilter { get; set; } = true;

        /// <summary>Calculates worksheet column widths from exported values.</summary>
        public bool AutoFit { get; set; }

        /// <summary>Uses the workbook's CellValue date and time number formats.</summary>
        public bool UseCellValueNumberFormats { get; set; }

        /// <summary>Writes explicit row and cell references. Disable for a smaller contiguous worksheet package.</summary>
        public bool IncludeCellReferences { get; set; } = true;

        /// <summary>Stores repeated text in the workbook shared-string table instead of writing inline strings.</summary>
        public bool UseSharedStrings { get; set; } = true;

        /// <summary>Excel date system used for temporal values.</summary>
        public ExcelDateSystem DateSystem { get; set; } = ExcelDateSystem.NineteenHundred;
    }
}
