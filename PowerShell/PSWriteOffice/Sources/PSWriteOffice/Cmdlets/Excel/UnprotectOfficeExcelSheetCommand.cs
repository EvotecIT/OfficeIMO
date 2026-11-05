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

/// <summary>Removes worksheet protection.</summary>
/// <example>
/// <summary>Remove sheet protection</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
///     Unprotect-OfficeExcelSheet -Sheet Inputs
/// }</code>
/// <para>Unprotects the Inputs sheet.</para>
/// </example>
/// <example>
/// <summary>Update a protected sheet and protect it again</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
///     ExcelSheet 'Rates' {
///         Unprotect-OfficeExcelSheet
///         ExcelCell B2 1.09
///         ExcelCell B3 0.86
///         Protect-OfficeExcelSheet -AllowSort
///     }
/// }</code>
/// <para>Unprotects the price sheet, writes the new rates and restores protection with sorting allowed.</para>
/// </example>
[Cmdlet(VerbsSecurity.Unprotect, "OfficeExcelSheet", DefaultParameterSetName = "Default")]
public sealed class UnprotectOfficeExcelSheetCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        sheet.Unprotect();
        if (PassThru) WriteObject(sheet);
    }

}
