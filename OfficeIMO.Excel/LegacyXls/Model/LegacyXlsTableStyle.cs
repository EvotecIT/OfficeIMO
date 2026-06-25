namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a user-defined BIFF TableStyle record and its elements.
    /// </summary>
    public sealed class LegacyXlsTableStyle {
        private readonly List<LegacyXlsTableStyleElement> _elements = new();

        /// <summary>
        /// Initializes a new instance of the <see cref="LegacyXlsTableStyle"/> class.
        /// </summary>
        public LegacyXlsTableStyle(
            string name,
            bool appliesToPivotTables,
            bool appliesToTables,
            uint declaredElementCount,
            ushort headerRecordType,
            ushort headerFlags,
            int recordOffset,
            ushort recordType,
            int payloadLength) {
            Name = name;
            AppliesToPivotTables = appliesToPivotTables;
            AppliesToTables = appliesToTables;
            DeclaredElementCount = declaredElementCount;
            HeaderRecordType = headerRecordType;
            HeaderFlags = headerFlags;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
        }

        /// <summary>Gets the user-defined table style name.</summary>
        public string Name { get; }

        /// <summary>Gets whether this style can be applied to PivotTable views.</summary>
        public bool AppliesToPivotTables { get; }

        /// <summary>Gets whether this style can be applied to tables.</summary>
        public bool AppliesToTables { get; }

        /// <summary>Gets the number of following TableStyleElement records declared by the style.</summary>
        public uint DeclaredElementCount { get; }

        /// <summary>Gets the parsed TableStyleElement records attached to this style.</summary>
        public IReadOnlyList<LegacyXlsTableStyleElement> Elements => _elements;

        /// <summary>Gets the FRT header record type stored inside the payload.</summary>
        public ushort HeaderRecordType { get; }

        /// <summary>Gets the FRT header flags stored inside the payload.</summary>
        public ushort HeaderFlags { get; }

        /// <summary>Gets the BIFF stream offset of the source record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the source BIFF record type.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the source BIFF payload length.</summary>
        public int PayloadLength { get; }

        internal void AddElement(LegacyXlsTableStyleElement element) {
            _elements.Add(element);
        }
    }
}
