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
        /// ActiveX control storage was discovered in the OLE compound container. Current imports expose this as preserved compound metadata.
        /// </summary>
        ActiveXControl,

        /// <summary>
        /// Embedded package payload storage or streams were discovered in the OLE compound container. Current imports expose this as preserved compound metadata.
        /// </summary>
        EmbeddedPackage,

        /// <summary>
        /// Binary payload stream was discovered in the OLE compound container. Current imports expose this as preserved compound metadata.
        /// </summary>
        BinaryData,

        /// <summary>
        /// Fast-save state was discovered in the FIB.
        /// </summary>
        FastSave,

        /// <summary>
        /// Picture payloads were indicated by the FIB. Current imports expose this as preserved metadata.
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
        /// Comment or annotation story text exists but does not have readable projection tables.
        /// </summary>
        Comment,

        /// <summary>
        /// Revision tracking state exists but cannot be projected into the current document shape.
        /// </summary>
        RevisionTracking,

        /// <summary>
        /// Text box story text exists but is not projected yet.
        /// </summary>
        TextBox,

        /// <summary>
        /// Bookmark structures exist but are outside the currently projected simple body shape.
        /// </summary>
        Bookmark,

        /// <summary>
        /// Invalid or conflicting merged table cell descriptors were found.
        /// </summary>
        MergedTableCell,

        /// <summary>
        /// Nested table descriptors were found before nested table projection is supported.
        /// </summary>
        NestedTable,

        /// <summary>
        /// A legacy OLE document property exists but is not projected into Open XML properties.
        /// </summary>
        DocumentProperty
    }
}
