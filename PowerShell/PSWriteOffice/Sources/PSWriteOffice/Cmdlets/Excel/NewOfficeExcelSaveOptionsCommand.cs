using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates native workbook save settings.</summary>
/// <example><summary>Validate and recalculate a workbook on save.</summary>
/// <code>$options = New-OfficeExcelSaveOptions -Preflight -EvaluateFormulas -FullCalculationOnOpen -ValidateOpenXml
/// Save-OfficeExcel -Document $workbook -SaveOptions $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeExcelSaveOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(ExcelSaveOptions))]
public sealed class NewOfficeExcelSaveOptionsCommand : PSCmdlet {
    /// <summary>Run native worksheet preflight cleanup before saving.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Preflight { get; set; }
    /// <summary>Repair common defined-name issues before saving.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter RepairDefinedNames { get; set; }
    /// <summary>Validate the saved package against the Open XML schema.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ValidateOpenXml { get; set; }
    /// <summary>Use the standard package finalization path.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoFastPackageWriter { get; set; }
    /// <summary>Evaluate supported formulas and write their cached results.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter EvaluateFormulas { get; set; }
    /// <summary>Remove cached formula results unless evaluation is requested.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ClearCachedFormulaResults { get; set; }
    /// <summary>Mark formulas dirty for the next application open.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter MarkFormulasDirty { get; set; }
    /// <summary>Request a full recalculation when the workbook is opened.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter FullCalculationOnOpen { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(ExcelDocumentService.CreateSaveOptions(
        Preflight.IsPresent, RepairDefinedNames.IsPresent, ValidateOpenXml.IsPresent,
        NoFastPackageWriter.IsPresent, EvaluateFormulas.IsPresent, ClearCachedFormulaResults.IsPresent,
        MarkFormulasDirty.IsPresent, FullCalculationOnOpen.IsPresent) ?? new ExcelSaveOptions());
}
