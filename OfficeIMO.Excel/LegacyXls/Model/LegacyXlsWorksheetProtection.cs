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
        /// <param name="permissions">Optional enhanced worksheet-protection permission exceptions.</param>
        public LegacyXlsWorksheetProtection(
            bool isProtected,
            string? legacyPasswordHash = null,
            bool? protectObjects = null,
            bool? protectScenarios = null,
            LegacyXlsWorksheetProtectionPermissions? permissions = null) {
            IsProtected = isProtected;
            LegacyPasswordHash = legacyPasswordHash;
            ProtectObjects = protectObjects;
            ProtectScenarios = protectScenarios;
            Permissions = permissions;
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

        /// <summary>
        /// Gets enhanced worksheet-protection permission exceptions parsed from BIFF8 FeatHdr metadata.
        /// </summary>
        public LegacyXlsWorksheetProtectionPermissions? Permissions { get; }

        internal LegacyXlsWorksheetProtection WithLegacyPasswordHash(ushort passwordHash) {
            return new LegacyXlsWorksheetProtection(IsProtected, passwordHash.ToString("X4"), ProtectObjects, ProtectScenarios, Permissions);
        }

        internal LegacyXlsWorksheetProtection WithObjectProtection(bool isProtected) {
            return new LegacyXlsWorksheetProtection(IsProtected, LegacyPasswordHash, isProtected, ProtectScenarios, Permissions);
        }

        internal LegacyXlsWorksheetProtection WithScenarioProtection(bool isProtected) {
            return new LegacyXlsWorksheetProtection(IsProtected, LegacyPasswordHash, ProtectObjects, isProtected, Permissions);
        }

        internal LegacyXlsWorksheetProtection WithPermissions(LegacyXlsWorksheetProtectionPermissions permissions) {
            return new LegacyXlsWorksheetProtection(IsProtected, LegacyPasswordHash, ProtectObjects, ProtectScenarios, permissions);
        }
    }
}
