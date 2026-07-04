namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>
    /// Identifies preserve-only compound storage categories discovered during legacy DOC import.
    /// </summary>
    public enum LegacyDocCompoundFeatureKind {
        /// <summary>
        /// VBA project storage was discovered in the OLE compound container.
        /// </summary>
        VbaProject,

        /// <summary>
        /// Embedded OLE object storage was discovered in the OLE compound container.
        /// </summary>
        OleObject,

        /// <summary>
        /// ActiveX control storage was discovered in the OLE compound container.
        /// </summary>
        ActiveXControl,

        /// <summary>
        /// Embedded package payload storage or streams were discovered in the OLE compound container.
        /// </summary>
        EmbeddedPackage,

        /// <summary>
        /// Binary payload stream was discovered in the OLE compound container.
        /// </summary>
        BinaryData
    }
}
