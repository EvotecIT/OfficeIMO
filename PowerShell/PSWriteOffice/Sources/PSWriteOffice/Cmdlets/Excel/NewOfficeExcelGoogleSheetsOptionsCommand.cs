using System;
using System.Collections;
using System.Management.Automation;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates native Excel-to-Google-Sheets planning and export settings without contacting Google.</summary>
/// <example><summary>Review a workbook with explicit formula and chart policies before uploading it.</summary>
/// <code>$options = New-OfficeExcelGoogleSheetsOptions -Title 'Quarterly forecast' -UnsupportedFormula PreserveWithWarning
/// Export-OfficeExcelGoogleSpreadsheet -Document $workbook -Options $options -PlanOnly</code></example>
[Cmdlet(VerbsCommon.New, "OfficeExcelGoogleSheetsOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(GoogleSheetsSaveOptions))]
public sealed class NewOfficeExcelGoogleSheetsOptionsCommand : OfficeCmdlet {
    /// <summary>Title of a new spreadsheet. Omission retains the native default.</summary>
    [Parameter(ParameterSetName = "Default"), ValidateNotNullOrEmpty]
    public string? Title { get; set; }
    /// <summary>Drive destination folder identifier. Omission retains the native default.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? FolderId { get; set; }
    /// <summary>Shared-drive identifier. Omission retains the native default.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? DriveId { get; set; }
    /// <summary>Existing spreadsheet identifier for replacement. Omission retains the native default.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? SpreadsheetId { get; set; }
    /// <summary>Drive version observed before replacement; the native exporter checks for changes. Omission retains the native default.</summary>
    [Parameter]
    public long? ExpectedDriveVersion { get; set; }
    /// <summary>Policy for formulas that cannot be translated. Omission retains the native default.</summary>
    [Parameter]
    public GoogleSheetsUnsupportedFormulaMode UnsupportedFormula { get; set; }
    /// <summary>Policy for charts without a reliable native equivalent. Omission retains the native default.</summary>
    [Parameter]
    public UnsupportedFeatureMode Charts { get; set; }
    /// <summary>Policy for unsupported pivot tables. Omission retains the native default.</summary>
    [Parameter]
    public UnsupportedFeatureMode PivotTables { get; set; }
    /// <summary>Policy for print layout without an equivalent. Omission retains the native default.</summary>
    [Parameter]
    public UnsupportedFeatureMode PrintLayout { get; set; }
    /// <summary>Target spreadsheet locale. Omission retains the native default.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Locale { get; set; }
    /// <summary>Target spreadsheet time zone. Omission retains the native default.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? TimeZone { get; set; }
    /// <summary>When the target recalculates formulas. Omission retains the native default.</summary>
    [Parameter]
    public GoogleSheetsRecalculationInterval RecalculationInterval { get; set; }
    /// <summary>Function-name replacements; keys and values are native formula function names.</summary>
    [Parameter]
    public Hashtable? FunctionMapping { get; set; }
    /// <summary>Editor email addresses for translated protected ranges.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string[]? Editor { get; set; }
    /// <summary>Whether domain users may edit protected content. Omission preserves the native policy.</summary>
    [Parameter]
    public SwitchParameter DomainUsersCanEdit { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new GoogleSheetsSaveOptions();
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Title))) options.Title = Title;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(FolderId))) options.Location.FolderId = FolderId;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(DriveId))) options.Location.DriveId = DriveId;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(SpreadsheetId))) options.Location.ExistingFileId = SpreadsheetId;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ExpectedDriveVersion))) options.Replace.ExpectedDriveVersion = ExpectedDriveVersion;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(UnsupportedFormula))) options.Formulas.UnsupportedFormulaMode = UnsupportedFormula;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Charts))) options.UnsupportedFeatures.Charts = Charts;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(PivotTables))) options.UnsupportedFeatures.PivotTables = PivotTables;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(PrintLayout))) options.UnsupportedFeatures.PrintLayout = PrintLayout;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Locale))) options.Spreadsheet.Locale = Locale;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TimeZone))) options.Spreadsheet.TimeZone = TimeZone;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(RecalculationInterval))) options.Spreadsheet.RecalculationInterval = RecalculationInterval;
        if (FunctionMapping != null) {
            foreach (DictionaryEntry item in FunctionMapping) {
                if (item.Key is not string name || string.IsNullOrWhiteSpace(name) ||
                    item.Value is not string replacement || string.IsNullOrWhiteSpace(replacement))
                    throw new PSArgumentException("FunctionMapping must contain nonempty function-name strings.");
                options.Formulas.FunctionMappings.Add(name, replacement);
            }
        }
        if (Editor != null) foreach (string address in Editor) options.Protection.EditorEmailAddresses.Add(address);
        if (BoundSwitch(nameof(DomainUsersCanEdit)) is bool domainUsersCanEdit)
            options.Protection.DomainUsersCanEdit = domainUsersCanEdit;
        WriteObject(options);
    }
}
