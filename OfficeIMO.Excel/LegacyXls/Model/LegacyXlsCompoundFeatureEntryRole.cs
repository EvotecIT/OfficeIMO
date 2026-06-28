namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Classifies preserve-only compound directory entries discovered in a legacy XLS OLE container.
    /// </summary>
    public enum LegacyXlsCompoundFeatureEntryRole {
        /// <summary>The entry role could not be classified more specifically.</summary>
        Unknown,

        /// <summary>The entry is the root VBA project storage.</summary>
        VbaProjectStorage,

        /// <summary>The entry is the VBA storage under the project.</summary>
        VbaStorage,

        /// <summary>The entry is the VBA dir stream.</summary>
        VbaDirStream,

        /// <summary>The entry is a VBA project metadata stream.</summary>
        VbaProjectStream,

        /// <summary>The entry is a VBA module, sheet, class, or ThisWorkbook stream.</summary>
        VbaModuleStream,

        /// <summary>The entry is the root OLE object pool storage.</summary>
        OleObjectPoolStorage,

        /// <summary>The entry is an OLE native payload stream.</summary>
        OleNativeStream,

        /// <summary>The entry is an OLE payload stream.</summary>
        OleStream,

        /// <summary>The entry is an embedded OLE object storage.</summary>
        OleObjectStorage,

        /// <summary>The entry is a CryptoAPI digital signature stream.</summary>
        DigitalSignatureStream,

        /// <summary>The entry is a CryptoAPI digital signature storage.</summary>
        DigitalSignatureStorage,

        /// <summary>The entry is an XML digital signature storage.</summary>
        XmlDigitalSignatureStorage,

        /// <summary>The entry is an XML digital signature stream.</summary>
        XmlDigitalSignatureStream
    }
}
