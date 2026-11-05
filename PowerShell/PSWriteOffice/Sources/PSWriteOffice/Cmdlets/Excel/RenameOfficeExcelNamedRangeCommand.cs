using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Renames a defined name.</summary>
/// <example>
/// <summary>Rename a defined name</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
///     Rename-OfficeExcelNamedRange Rate TaxRate
/// }</code>
/// <para>Renames Rate to TaxRate; formulas that use the name follow the change.</para>
/// </example>
/// <example>
/// <summary>Apply a naming convention</summary>
/// <code>$names = Get-OfficeExcelNamedRange -Path .\Finance\Pricing-Model.xlsx | Where-Object SheetName -eq 'Inputs'
/// Open-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
///     foreach ($name in $names) {
///         Rename-OfficeExcelNamedRange -Name $name.Name -NewName "in_$($name.Name)" -ValidationMode Strict
///     }
/// }</code>
/// <para>Prefixes every input name of the pricing model with in_ so inputs are easy to tell apart from calculations.</para>
/// </example>
[Cmdlet(VerbsCommon.Rename, "OfficeExcelNamedRange", DefaultParameterSetName = "Default")]
public sealed class RenameOfficeExcelNamedRangeCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Named range name.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public string Name { get; set; } = null!;

    /// <summary>New named range name.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 1)]
    public string NewName { get; set; } = null!;

    /// <summary>Use workbook-global scope from inside the DSL.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Global { get; set; }

    /// <summary>Defined-name validation mode.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelDefinedNameValidationMode? ValidationMode { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ExcelCommandTargets.Document(Document);
        ExcelSheet? scope = Global ? null : Sheet == null ? (Document == null ? ExcelDslContext.Current?.CurrentSheet : null) : Sheet.Resolve(document.Sheets, sheet => sheet.Name);
        bool changed = document.RenameNamedRange(Name, NewName, scope, save: false, validationMode: ValidationMode ?? ExcelDefinedNameValidationMode.Sanitize);
        if (PassThru) WriteObject(changed);
    }

}
