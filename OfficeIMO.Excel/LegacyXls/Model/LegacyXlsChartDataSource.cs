namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes shallow metadata decoded from a chart BRAI data-source record.
    /// </summary>
    public sealed class LegacyXlsChartDataSource {
        internal LegacyXlsChartDataSource(
            byte sourceId,
            string sourceIdName,
            byte referenceType,
            string referenceTypeName,
            ushort flags,
            bool usesCustomNumberFormat,
            ushort numberFormatId,
            ushort formulaByteCount,
            int formulaBytesAvailable,
            bool formulaByteCountFitsPayload) {
            SourceId = sourceId;
            SourceIdName = sourceIdName ?? throw new ArgumentNullException(nameof(sourceIdName));
            ReferenceType = referenceType;
            ReferenceTypeName = referenceTypeName ?? throw new ArgumentNullException(nameof(referenceTypeName));
            Flags = flags;
            UsesCustomNumberFormat = usesCustomNumberFormat;
            NumberFormatId = numberFormatId;
            FormulaByteCount = formulaByteCount;
            FormulaBytesAvailable = formulaBytesAvailable;
            FormulaByteCountFitsPayload = formulaByteCountFitsPayload;
        }

        /// <summary>Gets the raw BRAI source identifier.</summary>
        public byte SourceId { get; }

        /// <summary>Gets the decoded BRAI source identifier name.</summary>
        public string SourceIdName { get; }

        /// <summary>Gets the raw BRAI reference type.</summary>
        public byte ReferenceType { get; }

        /// <summary>Gets the decoded BRAI reference type name.</summary>
        public string ReferenceTypeName { get; }

        /// <summary>Gets the raw BRAI flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets a value indicating whether the record uses its own number format id.</summary>
        public bool UsesCustomNumberFormat { get; }

        /// <summary>Gets the raw number format identifier.</summary>
        public ushort NumberFormatId { get; }

        /// <summary>Gets the declared ChartParsedFormula byte count.</summary>
        public ushort FormulaByteCount { get; }

        /// <summary>Gets the number of formula bytes physically present after the BRAI fixed fields.</summary>
        public int FormulaBytesAvailable { get; }

        /// <summary>Gets a value indicating whether the declared formula byte count fits the available payload.</summary>
        public bool FormulaByteCountFitsPayload { get; }
    }
}
