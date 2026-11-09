using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates or updates a defined name.</summary>
/// <example>
/// <summary>Name an input cell</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
///     ExcelSheet 'Assumptions' {
///         ExcelNamedRange TaxRate 'B2'
///     }
/// }</code>
/// <para>Names B2 on the Assumptions sheet TaxRate so formulas can use =Amount*TaxRate.</para>
/// </example>
/// <example>
/// <summary>Define the inputs of a model</summary>
/// <code>New-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
///     ExcelSheet 'Inputs' {
///         ExcelCell A1 'Exchange rate EUR/USD'
///         ExcelCell B1 1.08
///         ExcelCell A2 'Target margin'
///         ExcelCell B2 0.32 -NumberFormat '0%'
///         ExcelNamedRange EurUsd 'B1'
///         ExcelNamedRange TargetMargin 'B2'
///     }
///     ExcelSheet 'Prices' {
///         ExcelCell A1 'List price (USD)'
///         ExcelCell B1 -Formula '=1000*EurUsd/(1-TargetMargin)'
///     }
/// }</code>
/// <para>Writes the model inputs and names each one, so the calculation sheets read like the business rules they implement.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelNamedRange", DefaultParameterSetName = "Default")]
[Alias("ExcelNamedRange")]
public sealed class SetOfficeExcelNamedRangeCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Name of the defined range.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public string Name { get; set; } = null!;

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 1)]
    public string Range { get; set; } = null!;

    /// <summary>Mark the defined name as hidden.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Hidden { get; set; }

    /// <summary>Validate or sanitize the defined name.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelDefinedNameValidationMode? ValidationMode { get; set; }

    /// <summary>Force a workbook-global name even inside a sheet block.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Global { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ExcelCommandTargets.Document(Document);
        ExcelSheet? scope = Global ? null : Sheet == null ? (Document == null ? ExcelDslContext.Current?.CurrentSheet : null) : Sheet.Resolve(document.Sheets, sheet => sheet.Name);
        document.SetNamedRange(Name, Range, scope, save: false, hidden: Hidden, validationMode: ValidationMode ?? ExcelDefinedNameValidationMode.Sanitize);
        if (PassThru) {
            string name = ExcelDocument.NormalizeDefinedName(Name, ValidationMode ?? ExcelDefinedNameValidationMode.Sanitize);
            WriteObject(document.ListNamedRanges().Single(item => item.Name == name && item.SheetName == scope?.Name));
        }
    }

}
