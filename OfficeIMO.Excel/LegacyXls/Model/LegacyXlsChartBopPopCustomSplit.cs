namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded BopPopCustom secondary bar/pie membership bits.
    /// </summary>
    public sealed class LegacyXlsChartBopPopCustomSplit {
        internal LegacyXlsChartBopPopCustomSplit(ushort bitCount, int bitmapBytesAvailable, IReadOnlyList<int> secondaryDataPointIndexes, bool noSecondaryDataPointsMarker) {
            BitCount = bitCount;
            BitmapBytesAvailable = bitmapBytesAvailable;
            SecondaryDataPointIndexes = secondaryDataPointIndexes ?? throw new ArgumentNullException(nameof(secondaryDataPointIndexes));
            NoSecondaryDataPointsMarker = noSecondaryDataPointsMarker;
        }

        /// <summary>Gets the declared BopPopCustom bit count, including the final additional bit.</summary>
        public ushort BitCount { get; }

        /// <summary>Gets the number of data point membership bits declared by the record.</summary>
        public int DataPointCount => BitCount == 0 ? 0 : BitCount - 1;

        /// <summary>Gets the expected custom-split bitmap byte count.</summary>
        public int ExpectedBitmapByteCount => 1 + (BitCount / 8);

        /// <summary>Gets the number of bitmap bytes available in the record payload.</summary>
        public int BitmapBytesAvailable { get; }

        /// <summary>Gets whether the declared bitmap is fully present in the payload.</summary>
        public bool HasCompleteBitmap => BitmapBytesAvailable >= ExpectedBitmapByteCount;

        /// <summary>Gets zero-based data point indexes assigned to the secondary bar/pie.</summary>
        public IReadOnlyList<int> SecondaryDataPointIndexes { get; }

        /// <summary>Gets whether the final additional bit marks an empty secondary bar/pie.</summary>
        public bool NoSecondaryDataPointsMarker { get; }

        /// <summary>Gets whether the empty-secondary marker is consistent with the decoded data point bits.</summary>
        public bool HasConsistentNoSecondaryDataPointsMarker => !NoSecondaryDataPointsMarker || SecondaryDataPointIndexes.Count == 0;
    }
}
