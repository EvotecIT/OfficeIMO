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

/// <summary>Sets the workbook date system (1900 or 1904).</summary>
/// <example>
/// <summary>Use the 1904 date system</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Legacy-Ledger.xlsx {
///     ExcelDateSystem 1904
/// }</code>
/// <para>Switches the workbook to the 1904 date system used by some older Mac workbooks.</para>
/// </example>
/// <example>
/// <summary>Match a partner workbook before copying dates</summary>
/// <code>New-OfficeExcel -Path .\Finance\Consolidated.xlsx {
///     ExcelDateSystem 1904
///     Copy-OfficeExcelSheet -Source .\Partners\Ledger-1904.xlsx -Sheet Ledger
/// }</code>
/// <para>Sets the date system of a new workbook to match a 1904-based partner file before copying its sheets in.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelDateSystem", DefaultParameterSetName = "Default")]
[Alias("ExcelDateSystem")]
public sealed class SetOfficeExcelDateSystemCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Date system to use for Excel date serials.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public ExcelDateSystem DateSystem { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ExcelCommandTargets.Document(Document);
        document.DateSystem = DateSystem;
        if (PassThru) WriteObject(document);
    }

}
