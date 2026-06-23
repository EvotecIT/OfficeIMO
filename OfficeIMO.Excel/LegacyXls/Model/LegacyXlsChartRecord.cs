namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a preserve-only chart BIFF record discovered during legacy XLS import.
    /// </summary>
    public sealed class LegacyXlsChartRecord {
        /// <summary>
        /// Creates chart BIFF record metadata.
        /// </summary>
        public LegacyXlsChartRecord(
            LegacyXlsChartRecordKind kind,
            string recordName,
            string? sheetName,
            int recordOffset,
            ushort recordType,
            int payloadLength,
            string? chartTypeName = null,
            int? chartX = null,
            int? chartY = null,
            int? chartWidth = null,
            int? chartHeight = null,
            ushort? axisType = null,
            string? axisTypeName = null) {
            if (payloadLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(payloadLength));
            }

            Kind = kind;
            RecordName = recordName ?? throw new ArgumentNullException(nameof(recordName));
            SheetName = sheetName;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
            ChartTypeName = string.IsNullOrWhiteSpace(chartTypeName) ? null : chartTypeName;
            ChartX = chartX;
            ChartY = chartY;
            ChartWidth = chartWidth;
            ChartHeight = chartHeight;
            AxisType = axisType;
            AxisTypeName = string.IsNullOrWhiteSpace(axisTypeName) ? null : axisTypeName;
        }

        /// <summary>Gets the shallow chart record category.</summary>
        public LegacyXlsChartRecordKind Kind { get; }

        /// <summary>Gets the BIFF record name.</summary>
        public string RecordName { get; }

        /// <summary>Gets the worksheet or chart sheet name associated with the record, when known.</summary>
        public string? SheetName { get; }

        /// <summary>Gets the byte offset of the BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }

        /// <summary>Gets the decoded chart family name for BIFF chart-type records, when available.</summary>
        public string? ChartTypeName { get; }

        /// <summary>Gets the decoded chart rectangle X coordinate from Chart records, when present.</summary>
        public int? ChartX { get; }

        /// <summary>Gets the decoded chart rectangle Y coordinate from Chart records, when present.</summary>
        public int? ChartY { get; }

        /// <summary>Gets the decoded chart rectangle width from Chart records, when present.</summary>
        public int? ChartWidth { get; }

        /// <summary>Gets the decoded chart rectangle height from Chart records, when present.</summary>
        public int? ChartHeight { get; }

        /// <summary>Gets the raw axis type from Axis records, when present.</summary>
        public ushort? AxisType { get; }

        /// <summary>Gets the decoded axis type name from Axis records, when present.</summary>
        public string? AxisTypeName { get; }
    }
}
