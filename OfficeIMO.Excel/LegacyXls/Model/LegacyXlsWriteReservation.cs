namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents workbook write-reservation metadata parsed from a BIFF FileSharing record.
    /// </summary>
    public sealed class LegacyXlsWriteReservation {
        /// <summary>
        /// Creates workbook write-reservation metadata.
        /// </summary>
        /// <param name="readOnlyRecommended">Whether applications should recommend opening the workbook as read-only.</param>
        /// <param name="legacyPasswordHash">Optional 16-bit write-reservation password verifier formatted as four uppercase hexadecimal digits.</param>
        /// <param name="userName">Optional user name associated with the write reservation.</param>
        public LegacyXlsWriteReservation(bool readOnlyRecommended, string? legacyPasswordHash = null, string? userName = null) {
            ReadOnlyRecommended = readOnlyRecommended;
            LegacyPasswordHash = legacyPasswordHash;
            UserName = string.IsNullOrWhiteSpace(userName) ? null : userName;
        }

        /// <summary>
        /// Gets whether applications should recommend opening the workbook as read-only.
        /// </summary>
        public bool ReadOnlyRecommended { get; }

        /// <summary>
        /// Gets the optional legacy write-reservation password verifier.
        /// </summary>
        public string? LegacyPasswordHash { get; }

        /// <summary>
        /// Gets the optional user name associated with the write reservation.
        /// </summary>
        public string? UserName { get; }
    }
}
