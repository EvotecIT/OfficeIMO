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
        /// <param name="protectObjects">Whether worksheet drawing objects are protected, when declared by ObjProtect.</param>
        /// <param name="protectScenarios">Whether worksheet scenarios are protected, when declared by ScenarioProtect.</param>
        public LegacyXlsWorksheetProtection(
            bool isProtected,
            string? legacyPasswordHash = null,
            bool? protectObjects = null,
            bool? protectScenarios = null) {
            IsProtected = isProtected;
            LegacyPasswordHash = legacyPasswordHash;
            ProtectObjects = protectObjects;
            ProtectScenarios = protectScenarios;
        }

        /// <summary>
        /// Gets whether worksheet protection is enabled.
        /// </summary>
        public bool IsProtected { get; }

        /// <summary>
        /// Gets the optional legacy worksheet protection password verifier.
        /// </summary>
        public string? LegacyPasswordHash { get; }

        /// <summary>
        /// Gets whether worksheet drawing objects are protected, when the legacy record declares it.
        /// </summary>
        public bool? ProtectObjects { get; }

        /// <summary>
        /// Gets whether worksheet scenarios are protected, when the legacy record declares it.
        /// </summary>
        public bool? ProtectScenarios { get; }

        internal LegacyXlsWorksheetProtection WithLegacyPasswordHash(ushort passwordHash) {
            return new LegacyXlsWorksheetProtection(IsProtected, passwordHash.ToString("X4"), ProtectObjects, ProtectScenarios);
        }

        internal LegacyXlsWorksheetProtection WithObjectProtection(bool isProtected) {
            return new LegacyXlsWorksheetProtection(IsProtected, LegacyPasswordHash, isProtected, ProtectScenarios);
        }

        internal LegacyXlsWorksheetProtection WithScenarioProtection(bool isProtected) {
            return new LegacyXlsWorksheetProtection(IsProtected, LegacyPasswordHash, ProtectObjects, isProtected);
        }
    }
}
