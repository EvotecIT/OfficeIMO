namespace OfficeIMO.Excel.Xlsb {
    /// <summary>
    /// Controls resource limits and unsupported-content reporting while importing XLSB workbooks.
    /// </summary>
    public sealed class XlsbImportOptions {
        /// <summary>Gets or sets the largest accepted decompressed package part. The default is 128 MiB.</summary>
        public int MaxPartBytes { get; set; } = 128 * 1024 * 1024;

        /// <summary>Gets or sets the aggregate decompressed package budget. The default is 512 MiB.</summary>
        public long MaxPackageBytes { get; set; } = 512L * 1024 * 1024;

        /// <summary>Gets or sets the largest accepted BIFF12 record payload. The default is 64 MiB.</summary>
        public int MaxRecordBytes { get; set; } = 64 * 1024 * 1024;

        /// <summary>Gets or sets the aggregate BIFF12 record limit across one workbook import.</summary>
        public int MaxRecordCount { get; set; } = 1_000_000;

        /// <summary>Gets or sets the maximum number of worksheets accepted from one workbook.</summary>
        public int MaxWorksheets { get; set; } = 16_384;

        /// <summary>Gets or sets the maximum number of populated cells projected from one workbook.</summary>
        public int MaxCells { get; set; } = 4_000_000;

        /// <summary>Gets or sets the maximum number of row metadata definitions projected from one workbook.</summary>
        public int MaxRowDefinitions { get; set; } = 100_000;

        /// <summary>Gets or sets the maximum number of shared-string items accepted.</summary>
        public int MaxSharedStrings { get; set; } = 1_000_000;

        /// <summary>Gets or sets the maximum number of merged ranges accepted from one workbook.</summary>
        public int MaxMergedRanges { get; set; } = 1_048_576;

        /// <summary>Gets or sets the maximum number of worksheet hyperlinks accepted from one workbook.</summary>
        public int MaxHyperlinks { get; set; } = 1_000_000;

        /// <summary>Gets or sets the maximum number of UTF-16 characters accepted in one BIFF12 string.</summary>
        public int MaxStringCharacters { get; set; } = 1_048_576;

        /// <summary>Gets or sets whether unknown records are listed in the import preservation report.</summary>
        public bool ReportPreservedRecords { get; set; } = true;

        internal void Validate() {
            if (MaxPartBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPartBytes));
            if (MaxPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPackageBytes));
            if (MaxRecordBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxRecordBytes));
            if (MaxRecordCount <= 0) throw new ArgumentOutOfRangeException(nameof(MaxRecordCount));
            if (MaxWorksheets <= 0) throw new ArgumentOutOfRangeException(nameof(MaxWorksheets));
            if (MaxCells <= 0) throw new ArgumentOutOfRangeException(nameof(MaxCells));
            if (MaxRowDefinitions <= 0) throw new ArgumentOutOfRangeException(nameof(MaxRowDefinitions));
            if (MaxSharedStrings <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSharedStrings));
            if (MaxMergedRanges <= 0) throw new ArgumentOutOfRangeException(nameof(MaxMergedRanges));
            if (MaxHyperlinks <= 0) throw new ArgumentOutOfRangeException(nameof(MaxHyperlinks));
            if (MaxStringCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxStringCharacters));
        }
    }
}
