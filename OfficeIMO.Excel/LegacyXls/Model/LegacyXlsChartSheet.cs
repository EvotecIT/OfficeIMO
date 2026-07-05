namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a legacy XLS chart sheet decoded from a BIFF chart-sheet substream.
    /// </summary>
    public sealed class LegacyXlsChartSheet {
        private readonly List<LegacyXlsChartSheetMetadataRecord> _metadataRecords = new();
        private readonly List<LegacyXlsSheetFutureMetadataRecord> _futureMetadataRecords = new();
        private readonly Dictionary<LegacyXlsChartRecordKind, int> _chartRecordsByKind = new();
        private readonly Dictionary<string, int> _chartRecordsByChartType = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Creates chart-sheet metadata.
        /// </summary>
        /// <param name="name">Sheet name from the BoundSheet8 record.</param>
        /// <param name="streamOffset">Byte offset of the chart-sheet substream in the BIFF workbook stream.</param>
        /// <param name="visibility">Legacy sheet visibility flag.</param>
        /// <param name="sheetType">Legacy BoundSheet8 sheet type flag.</param>
        public LegacyXlsChartSheet(
            string name,
            int streamOffset,
            byte visibility,
            byte sheetType) {
            Name = name;
            StreamOffset = streamOffset;
            Visibility = visibility;
            SheetType = sheetType;
        }

        /// <summary>
        /// Gets the chart-sheet name from the BoundSheet8 record.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the byte offset of the chart-sheet substream in the BIFF workbook stream.
        /// </summary>
        public int StreamOffset { get; }

        /// <summary>
        /// Gets the legacy visibility flag.
        /// </summary>
        public byte Visibility { get; }

        /// <summary>
        /// Gets the decoded sheet visibility state, when the BoundSheet value is recognized.
        /// </summary>
        public LegacyXlsSheetVisibility? VisibilityKind => LegacyXlsSheetVisibilityDecoder.ToKind(Visibility);

        /// <summary>
        /// Gets the decoded sheet visibility state name, or a hexadecimal fallback for unknown values.
        /// </summary>
        public string VisibilityName => LegacyXlsSheetVisibilityDecoder.ToName(Visibility);

        /// <summary>
        /// Gets the legacy BoundSheet8 sheet type flag.
        /// </summary>
        public byte SheetType { get; }

        /// <summary>
        /// Gets decoded metadata records from this chart-sheet substream.
        /// </summary>
        public IReadOnlyList<LegacyXlsChartSheetMetadataRecord> MetadataRecords => _metadataRecords;

        /// <summary>
        /// Gets extended metadata records decoded from this chart-sheet substream.
        /// </summary>
        public IReadOnlyList<LegacyXlsSheetFutureMetadataRecord> FutureMetadataRecords => _futureMetadataRecords;

        /// <summary>
        /// Gets the chart printed-size mode from a PrintSize record, when present.
        /// </summary>
        public ushort? ChartPrintSize { get; private set; }

        /// <summary>
        /// Gets the decoded chart printed-size mode, when the PrintSize value is recognized.
        /// </summary>
        public LegacyXlsChartPrintSize? ChartPrintSizeKind {
            get {
                switch (ChartPrintSize) {
                    case 0x0000: return LegacyXlsChartPrintSize.DefaultsUnchanged;
                    case 0x0001: return LegacyXlsChartPrintSize.FillPage;
                    case 0x0002: return LegacyXlsChartPrintSize.FitPage;
                    case 0x0003: return LegacyXlsChartPrintSize.DefinedInChartRecord;
                    default: return null;
                }
            }
        }

        /// <summary>
        /// Gets the decoded chart printed-size mode name, or a hexadecimal fallback for unknown values.
        /// </summary>
        public string? ChartPrintSizeName {
            get {
                if (!ChartPrintSize.HasValue) {
                    return null;
                }

                LegacyXlsChartPrintSize? kind = ChartPrintSizeKind;
                return kind.HasValue ? kind.Value.ToString() : $"PrintSize:0x{ChartPrintSize.Value:X4}";
            }
        }

        /// <summary>
        /// Gets the number of chart text object records seen in this chart-sheet substream.
        /// </summary>
        public int ChartTextObjectCount { get; private set; }

        /// <summary>
        /// Gets the number of supported chart records seen in this chart-sheet substream.
        /// </summary>
        public int ChartRecordCount { get; private set; }

        /// <summary>
        /// Gets chart records from this chart sheet grouped by shallow category.
        /// </summary>
        public IReadOnlyDictionary<LegacyXlsChartRecordKind, int> ChartRecordsByKind => _chartRecordsByKind;

        /// <summary>
        /// Gets chart type records from this chart sheet grouped by decoded chart family.
        /// </summary>
        public IReadOnlyDictionary<string, int> ChartRecordsByChartType => _chartRecordsByChartType;

        internal void AddMetadataRecord(LegacyXlsChartSheetMetadataKind kind, int recordOffset, ushort recordType) {
            _metadataRecords.Add(new LegacyXlsChartSheetMetadataRecord(kind, recordOffset, recordType));
        }

        internal void AddFutureMetadataRecord(LegacyXlsSheetFutureMetadataRecord record) {
            _futureMetadataRecords.Add(record);
            AddMetadataRecord(LegacyXlsChartSheetMetadataKind.FutureMetadata, record.RecordOffset, record.RecordType);
        }

        internal void AddChartRecord(LegacyXlsChartRecord record) {
            if (record == null) throw new ArgumentNullException(nameof(record));

            ChartRecordCount++;
            Increment(_chartRecordsByKind, record.Kind);
            if (!string.IsNullOrWhiteSpace(record.ChartTypeName)) {
                Increment(_chartRecordsByChartType, record.ChartTypeName!);
            }
        }

        internal void SetChartPrintSize(ushort value) {
            ChartPrintSize = value;
        }

        internal void IncrementChartTextObjectCount() {
            ChartTextObjectCount++;
        }

        private static void Increment<TKey>(Dictionary<TKey, int> counts, TKey key) where TKey : notnull {
            counts.TryGetValue(key, out int count);
            counts[key] = count + 1;
        }
    }
}
