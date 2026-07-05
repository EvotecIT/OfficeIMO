namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a projected worksheet table/list definition parsed from BIFF Feature11/Feature12 records.
    /// </summary>
    public sealed class LegacyXlsTableDefinition {
        /// <summary>
        /// Initializes a new instance of the <see cref="LegacyXlsTableDefinition"/> class.
        /// </summary>
        public LegacyXlsTableDefinition(
            string name,
            string range,
            bool hasHeaderRow,
            uint totalRowCount,
            bool hasAutoFilter,
            uint idList,
            int recordOffset,
            ushort recordType,
            int payloadLength) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Table name cannot be empty.", nameof(name));
            }

            if (string.IsNullOrWhiteSpace(range)) {
                throw new ArgumentException("Table range cannot be empty.", nameof(range));
            }

            Name = name;
            Range = range;
            HasHeaderRow = hasHeaderRow;
            TotalRowCount = totalRowCount;
            HasAutoFilter = hasAutoFilter;
            IdList = idList;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
        }

        /// <summary>Gets the legacy table name.</summary>
        public string Name { get; }

        /// <summary>Gets the projected A1 table range.</summary>
        public string Range { get; }

        /// <summary>Gets the table style name parsed from a List12 table-style record, when present.</summary>
        public string? StyleName { get; private set; }

        /// <summary>Gets the table display name parsed from a List12 display-name record, when present.</summary>
        public string? DisplayName { get; private set; }

        /// <summary>Gets the table comment parsed from a List12 display-name record, when present.</summary>
        public string? Comment { get; private set; }

        /// <summary>Gets block-level table region formatting parsed from a List12 block-level record, when present.</summary>
        public LegacyXlsTableBlockLevelFormatting? BlockLevelFormatting { get; private set; }

        /// <summary>Gets whether the table declares a header row.</summary>
        public bool HasHeaderRow { get; }

        /// <summary>Gets the number of totals rows declared by the table definition.</summary>
        public uint TotalRowCount { get; }

        /// <summary>Gets whether the table declares a totals row.</summary>
        public bool HasTotalsRow => TotalRowCount > 0;

        /// <summary>Gets whether the table declares an AutoFilter.</summary>
        public bool HasAutoFilter { get; }

        /// <summary>Gets whether the table style applies first-column formatting, when specified.</summary>
        public bool? ShowFirstColumn { get; private set; }

        /// <summary>Gets whether the table style applies last-column formatting, when specified.</summary>
        public bool? ShowLastColumn { get; private set; }

        /// <summary>Gets whether the table style applies row stripe formatting, when specified.</summary>
        public bool? ShowRowStripes { get; private set; }

        /// <summary>Gets whether the table style applies column stripe formatting, when specified.</summary>
        public bool? ShowColumnStripes { get; private set; }

        /// <summary>Gets the BIFF table identifier.</summary>
        public uint IdList { get; }

        /// <summary>Gets the BIFF record offset.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the BIFF payload length.</summary>
        public int PayloadLength { get; }

        internal void ApplyStyle(
            string styleName,
            bool showFirstColumn,
            bool showLastColumn,
            bool showRowStripes,
            bool showColumnStripes) {
            if (string.IsNullOrWhiteSpace(styleName)) {
                return;
            }

            StyleName = styleName;
            ShowFirstColumn = showFirstColumn;
            ShowLastColumn = showLastColumn;
            ShowRowStripes = showRowStripes;
            ShowColumnStripes = showColumnStripes;
        }

        internal void ApplyDisplayMetadata(string? displayName, string? comment) {
            if (!string.IsNullOrWhiteSpace(displayName)) {
                DisplayName = displayName;
            }

            if (!string.IsNullOrWhiteSpace(comment)) {
                Comment = comment;
            }
        }

        internal void ApplyBlockLevelFormatting(LegacyXlsTableBlockLevelFormatting formatting) {
            BlockLevelFormatting = formatting ?? throw new ArgumentNullException(nameof(formatting));
        }
    }
}
