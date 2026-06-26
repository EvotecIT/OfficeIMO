namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the BIFF body shape represented by an ExternName record.
    /// </summary>
    public enum LegacyXlsExternalNameBodyKind {
        /// <summary>The body shape could not be determined from the supporting link and flags.</summary>
        Unknown,

        /// <summary>The body describes an external defined name.</summary>
        ExternalDefinedName,

        /// <summary>The body describes an add-in user-defined function reference.</summary>
        AddInUdf,

        /// <summary>The body describes a DDE or OLE linked item with operation data.</summary>
        OleDdeLink,

        /// <summary>The body describes an OLE linked item.</summary>
        OleDataItem,

        /// <summary>The body describes a DDE linked item without operation data.</summary>
        DdeLinkNoOper
    }
}
