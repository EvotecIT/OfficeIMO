namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a legacy XLS sheet entry that is preserved as import metadata but not projected as a worksheet.
    /// </summary>
    public sealed class LegacyXlsUnsupportedSheet {
        private readonly List<LegacyXlsUnsupportedSheetMetadataRecord> _metadataRecords = new();
        private readonly List<LegacyXlsSheetFutureMetadataRecord> _futureMetadataRecords = new();
        private readonly Dictionary<LegacyXlsChartRecordKind, int> _chartRecordsByKind = new();
        private readonly Dictionary<string, int> _chartRecordsByChartType = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Creates unsupported legacy sheet metadata.
        /// </summary>
        /// <param name="name">Sheet name from the BoundSheet8 record.</param>
        /// <param name="streamOffset">Byte offset of the sheet substream in the BIFF workbook stream.</param>
        /// <param name="visibility">Legacy sheet visibility flag.</param>
        /// <param name="sheetType">Legacy BoundSheet8 sheet type flag.</param>
        /// <param name="kind">Unsupported sheet category.</param>
        public LegacyXlsUnsupportedSheet(
            string name,
            int streamOffset,
            byte visibility,
            byte sheetType,
            LegacyXlsUnsupportedSheetKind kind) {
            Name = name;
            StreamOffset = streamOffset;
            Visibility = visibility;
            SheetType = sheetType;
            Kind = kind;
        }

        /// <summary>
        /// Gets the sheet name from the BoundSheet8 record.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the byte offset of the sheet substream in the BIFF workbook stream.
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
        /// Gets the unsupported sheet category.
        /// </summary>
        public LegacyXlsUnsupportedSheetKind Kind { get; }

        /// <summary>
        /// Gets decoded metadata records from this unsupported sheet substream.
        /// </summary>
        public IReadOnlyList<LegacyXlsUnsupportedSheetMetadataRecord> MetadataRecords => _metadataRecords;

        /// <summary>
        /// Gets preserve-only extended metadata records decoded from this unsupported sheet substream.
        /// </summary>
        public IReadOnlyList<LegacyXlsSheetFutureMetadataRecord> FutureMetadataRecords => _futureMetadataRecords;

        /// <summary>
        /// Gets the chart printed-size mode from a PrintSize record, when this unsupported sheet is a chart sheet.
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
        /// Gets the number of chart text object records seen in this unsupported chart sheet substream.
        /// </summary>
        public int ChartTextObjectCount { get; private set; }

        /// <summary>
        /// Gets the number of preserve-only chart records seen in this unsupported chart sheet substream.
        /// </summary>
        public int ChartRecordCount { get; private set; }

        /// <summary>
        /// Gets preserve-only chart records from this unsupported chart sheet grouped by shallow category.
        /// </summary>
        public IReadOnlyDictionary<LegacyXlsChartRecordKind, int> ChartRecordsByKind => _chartRecordsByKind;

        /// <summary>
        /// Gets preserve-only chart type records from this unsupported chart sheet grouped by decoded chart family.
        /// </summary>
        public IReadOnlyDictionary<string, int> ChartRecordsByChartType => _chartRecordsByChartType;

        internal void AddMetadataRecord(LegacyXlsUnsupportedSheetMetadataKind kind, int recordOffset, ushort recordType) {
            _metadataRecords.Add(new LegacyXlsUnsupportedSheetMetadataRecord(kind, recordOffset, recordType));
        }

        internal void AddFutureMetadataRecord(LegacyXlsSheetFutureMetadataRecord record) {
            _futureMetadataRecords.Add(record);
            AddMetadataRecord(LegacyXlsUnsupportedSheetMetadataKind.FutureMetadata, record.RecordOffset, record.RecordType);
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
