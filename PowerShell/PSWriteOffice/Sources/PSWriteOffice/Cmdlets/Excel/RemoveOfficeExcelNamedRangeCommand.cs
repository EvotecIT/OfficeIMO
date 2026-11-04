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

/// <summary>Removes a defined name.</summary>
/// <example>
/// <summary>Remove a defined name</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
///     Remove-OfficeExcelNamedRange OldTaxRate
/// }</code>
/// <para>Deletes the OldTaxRate name from the workbook.</para>
/// </example>
/// <example>
/// <summary>Clean up names that no longer resolve</summary>
/// <code>$broken = Get-OfficeExcelNamedRange -Path .\Finance\Budget-2027.xlsx | Where-Object Reference -like '*#REF!*'
/// Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
///     foreach ($name in $broken) {
///         Remove-OfficeExcelNamedRange -Name $name.Name
///     }
/// }</code>
/// <para>Removes every defined name that points to #REF! and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficeExcelNamedRange", DefaultParameterSetName = "Default")]
public sealed class RemoveOfficeExcelNamedRangeCommand : PSCmdlet {
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

    /// <summary>Use workbook-global scope from inside the DSL.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Global { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ExcelCommandTargets.Document(Document);
        ExcelSheet? scope = Global ? null : Sheet == null ? (Document == null ? ExcelDslContext.Current?.CurrentSheet : null) : Sheet.Resolve(document.Sheets, sheet => sheet.Name);
        bool changed = document.RemoveNamedRange(Name, scope, save: false);
        if (PassThru) WriteObject(changed);
    }

}
