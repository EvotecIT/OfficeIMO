namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>The native worksheet visibility states accepted by Set-OfficeExcelSheet.</summary>
public enum ExcelSheetVisibility {
    /// <summary>Shows the worksheet.</summary>
    Visible,
    /// <summary>Hides the worksheet; spreadsheet applications can unhide it.</summary>
    Hidden,
    /// <summary>Hides the worksheet from the application's ordinary unhide dialog.</summary>
    VeryHidden
}
