namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>
    /// Identifies unsupported or preserve-only legacy DOC features discovered during import.
    /// </summary>
    public enum LegacyDocUnsupportedFeatureKind {
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
        BinaryData,

        /// <summary>
        /// Fast-save or quick-save state was discovered in the FIB.
        /// </summary>
        FastSave,

        /// <summary>
        /// Picture payloads were indicated by the FIB.
        /// </summary>
        Picture,

        /// <summary>
        /// Header or footer story text exists but is not projected yet.
        /// </summary>
        HeaderFooter,

        /// <summary>
        /// Multiple legacy DOC sections exist but are not projected yet.
        /// </summary>
        Section,

        /// <summary>
        /// Footnote story text exists but is not projected in the current document shape.
        /// </summary>
        Footnote,

        /// <summary>
        /// Endnote story text exists but is not projected yet.
        /// </summary>
        Endnote,

        /// <summary>
        /// Comment or annotation story text exists but is not projected yet.
        /// </summary>
        Comment,

        /// <summary>
        /// Revision tracking state exists but is not projected yet.
        /// </summary>
        RevisionTracking,

        /// <summary>
        /// Text box story text exists but is not projected yet.
        /// </summary>
        TextBox,

        /// <summary>
        /// Invalid or conflicting merged table cell descriptors were found.
        /// </summary>
        MergedTableCell
    }
}
