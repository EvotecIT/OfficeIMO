namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents worksheet protection metadata parsed from a legacy XLS worksheet substream.
    /// </summary>
    public sealed class LegacyXlsWorksheetProtection {
        /// <summary>
        /// Creates worksheet protection metadata.
        /// </summary>
        /// <param name="isProtected">Whether the worksheet protection flag is enabled.</param>
        /// <param name="legacyPasswordHash">Optional 16-bit legacy password verifier formatted as four uppercase hexadecimal digits.</param>
        public LegacyXlsWorksheetProtection(bool isProtected, string? legacyPasswordHash = null) {
            IsProtected = isProtected;
            LegacyPasswordHash = legacyPasswordHash;
        }

        /// <summary>
        /// Gets whether worksheet protection is enabled.
        /// </summary>
        public bool IsProtected { get; }

        /// <summary>
        /// Gets the optional legacy worksheet protection password verifier.
        /// </summary>
        public string? LegacyPasswordHash { get; }

        internal LegacyXlsWorksheetProtection WithLegacyPasswordHash(ushort passwordHash) {
            return new LegacyXlsWorksheetProtection(IsProtected, passwordHash.ToString("X4"));
        }
    }
}
