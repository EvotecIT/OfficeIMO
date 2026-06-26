namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes which axis component is formatted by a following LineFormat record.
    /// </summary>
    public sealed class LegacyXlsChartAxisLineFormat {
        internal LegacyXlsChartAxisLineFormat(ushort targetId, string targetName) {
            TargetId = targetId;
            TargetName = targetName ?? throw new ArgumentNullException(nameof(targetName));
        }

        /// <summary>Gets the raw AxisLine target identifier.</summary>
        public ushort TargetId { get; }

        /// <summary>Gets the decoded axis component name.</summary>
        public string TargetName { get; }
    }
}
