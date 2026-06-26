namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes preserve-only ChartFrtInfo metadata for chart future-record ranges.
    /// </summary>
    public sealed class LegacyXlsChartFutureRecordInfo {
        internal LegacyXlsChartFutureRecordInfo(byte originatorVersion, byte writerVersion, IReadOnlyList<LegacyXlsChartFutureRecordRange> ranges) {
            OriginatorVersion = originatorVersion;
            WriterVersion = writerVersion;
            Ranges = ranges?.ToArray() ?? Array.Empty<LegacyXlsChartFutureRecordRange>();
        }

        /// <summary>Gets the application version that originally created the chart future-record envelope.</summary>
        public byte OriginatorVersion { get; }

        /// <summary>Gets the application version that last saved the chart future-record envelope.</summary>
        public byte WriterVersion { get; }

        /// <summary>Gets the decoded originator version name.</summary>
        public string OriginatorVersionName => GetVersionName(OriginatorVersion);

        /// <summary>Gets the decoded writer version name.</summary>
        public string WriterVersionName => GetVersionName(WriterVersion);

        /// <summary>Gets the declared future-record identifier ranges.</summary>
        public IReadOnlyList<LegacyXlsChartFutureRecordRange> Ranges { get; }

        private static string GetVersionName(byte version) {
            return version switch {
                0x09 => "Version:0x09",
                0x0A => "Version:0x0A",
                0x0C => "Version:0x0C",
                0x0E => "Version:0x0E",
                0x0F => "Version:0x0F",
                _ => $"Unknown:0x{version:X2}"
            };
        }
    }
}
