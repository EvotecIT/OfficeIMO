namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies workbook-level BIFF metadata that has been decoded into the legacy import model.
    /// </summary>
    public enum LegacyXlsWorkbookMetadataKind {
        /// <summary>Backup-save preference from a Backup record.</summary>
        Backup,

        /// <summary>Workbook option flags from a BookBool record.</summary>
        BookOptions,

        /// <summary>Built-in function category count from a BuiltInFnGroupCount record.</summary>
        BuiltInFunctionGroupCount,

        /// <summary>Workbook text code page from a CodePage record.</summary>
        CodePage,

        /// <summary>VBA workbook object name from a CodeName record.</summary>
        CodeName,

        /// <summary>Workbook locale identifiers from a Country record.</summary>
        Country,

        /// <summary>Hidden-object display mode from a HideObj record.</summary>
        HiddenObjects,

        /// <summary>User interface code page from an InterfaceHdr record.</summary>
        InterfaceCodePage,

        /// <summary>End marker for the user interface record collection.</summary>
        InterfaceEnd,

        /// <summary>Natural language formula support flag from a UsesELFs record.</summary>
        NaturalLanguageFormulas,

        /// <summary>VBA project marker from an ObProj record.</summary>
        VbaProjectMarker,

        /// <summary>Reserved DSF record that must be ignored by importers.</summary>
        ReservedDsf,

        /// <summary>Opaque printer settings from a Pls record.</summary>
        PrinterSettings,

        /// <summary>Printed workbook sizing mode from a PrintSize record.</summary>
        PrintSize,

        /// <summary>Revision-tracking lock state from a Prot4Rev record.</summary>
        RevisionProtection,

        /// <summary>Revision-tracking password verifier from a Prot4RevPass record.</summary>
        RevisionProtectionPassword,

        /// <summary>Sheet tab identifier array from a TabId record.</summary>
        SheetTabIds,

        /// <summary>VBA project marker indicating the project contains no forms, modules, or class modules.</summary>
        VbaProjectNoMacrosMarker,

        /// <summary>Workbook window lock state from a WinProtect record.</summary>
        WindowProtection,

        /// <summary>Workbook window geometry and display flags from a Window1 record.</summary>
        Window,

        /// <summary>Last write user name from a WriteAccess record.</summary>
        WriteAccess
    }
}
