namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents workbook protection metadata parsed from the legacy XLS globals substream.
    /// </summary>
    public sealed class LegacyXlsWorkbookProtection {
        /// <summary>
        /// Creates workbook protection metadata.
        /// </summary>
        /// <param name="isProtected">Whether workbook protection is enabled.</param>
        /// <param name="legacyPasswordHash">Optional 16-bit legacy password verifier formatted as four uppercase hexadecimal digits.</param>
        public LegacyXlsWorkbookProtection(bool isProtected, string? legacyPasswordHash = null) {
            IsProtected = isProtected;
            LegacyPasswordHash = legacyPasswordHash;
        }

        /// <summary>
        /// Gets whether workbook protection is enabled.
        /// </summary>
        public bool IsProtected { get; }

        /// <summary>
        /// Gets the optional legacy workbook protection password verifier.
        /// </summary>
        public string? LegacyPasswordHash { get; }

        internal LegacyXlsWorkbookProtection WithLegacyPasswordHash(ushort passwordHash) {
            return new LegacyXlsWorkbookProtection(IsProtected, passwordHash.ToString("X4"));
        }
    }
}
