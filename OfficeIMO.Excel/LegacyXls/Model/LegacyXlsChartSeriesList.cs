namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes preserve-only SeriesList chart metadata.
    /// </summary>
    public sealed class LegacyXlsChartSeriesList {
        internal LegacyXlsChartSeriesList(ushort declaredSeriesCount, IReadOnlyList<ushort> seriesIndexes) {
            DeclaredSeriesCount = declaredSeriesCount;
            SeriesIndexes = seriesIndexes ?? throw new ArgumentNullException(nameof(seriesIndexes));
        }

        /// <summary>Gets the series index count declared by the record.</summary>
        public ushort DeclaredSeriesCount { get; }

        /// <summary>Gets decoded one-based Series record indexes.</summary>
        public IReadOnlyList<ushort> SeriesIndexes { get; }

        /// <summary>Gets the number of series indexes decoded from the available payload.</summary>
        public int DecodedSeriesCount => SeriesIndexes.Count;

        /// <summary>Gets whether the decoded index count matches the declared count.</summary>
        public bool HasCompleteSeriesIndexList => DeclaredSeriesCount == SeriesIndexes.Count;

        /// <summary>Gets whether every decoded series index is inside the BIFF chart series index range.</summary>
        public bool HasOnlyValidSeriesIndexes {
            get {
                foreach (ushort seriesIndex in SeriesIndexes) {
                    if (seriesIndex < 1 || seriesIndex > 0x00FE) {
                        return false;
                    }
                }

                return true;
            }
        }
    }
}
