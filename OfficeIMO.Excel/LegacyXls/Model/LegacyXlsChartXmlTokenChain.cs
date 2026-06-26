namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes preserve-only CrtMlFrt chart XML token-chain metadata.
    /// </summary>
    public sealed class LegacyXlsChartXmlTokenChain {
        internal LegacyXlsChartXmlTokenChain(uint declaredByteCount, int firstSegmentByteCount, uint trailingUnusedValue) {
            DeclaredByteCount = declaredByteCount;
            FirstSegmentByteCount = firstSegmentByteCount;
            TrailingUnusedValue = trailingUnusedValue;
        }

        /// <summary>Gets the declared XmlTkChain byte count, including optional continuation records.</summary>
        public uint DeclaredByteCount { get; }

        /// <summary>Gets the number of XmlTkChain bytes available in this CrtMlFrt record.</summary>
        public int FirstSegmentByteCount { get; }

        /// <summary>Gets the ignored trailing field value.</summary>
        public uint TrailingUnusedValue { get; }

        /// <summary>Gets whether the declared chain bytes fit in this record without continuation bytes.</summary>
        public bool IsCompleteInRecord => DeclaredByteCount <= (uint)FirstSegmentByteCount;

        /// <summary>Gets whether the ignored trailing field is zero.</summary>
        public bool HasZeroTrailingUnusedValue => TrailingUnusedValue == 0;
    }
}
