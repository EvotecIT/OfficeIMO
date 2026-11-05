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

/// <summary>Groups columns into a collapsible outline, or clears grouping.</summary>
/// <example>
/// <summary>Group monthly columns</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Budget-2027.xlsx {
///     ExcelSheet 'Departments' {
///         ExcelColumnGroup B M
///     }
/// }</code>
/// <para>Groups columns B to M, the monthly figures, next to the yearly total.</para>
/// </example>
/// <example>
/// <summary>Collapse quarters into their totals</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Forecast.xlsx {
///     ExcelSheet 'Forecast' {
///         ExcelColumnGroup B D -Collapsed
///         ExcelColumnGroup F H -Collapsed
///         ExcelColumnGroup J L -Collapsed
///         ExcelColumnGroup N P -Collapsed
///     }
/// }</code>
/// <para>Groups each quarter's three months and collapses them, so the sheet shows the quarterly totals first.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelColumnGroup", DefaultParameterSetName = "Default")]
[Alias("ExcelColumnGroup")]
public sealed class SetOfficeExcelColumnGroupCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>First 1-based column in the group.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public object StartColumn { get; set; } = null!;

    /// <summary>Last 1-based column in the group. Defaults to the start column.</summary>
    [Parameter(ParameterSetName = "Default", Position = 1)]
    public object? EndColumn { get; set; }

    /// <summary>Excel outline level from 1 through 7.</summary>
    [Parameter(ParameterSetName = "Default")]
    [ValidateRange(1, 7)]
    public int OutlineLevel { get; set; } = 1;

    /// <summary>Hide the grouped columns and mark the following summary column as collapsed.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Collapsed { get; set; }

    /// <summary>Hide the grouped columns without marking the group collapsed.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Hidden { get; set; }

    /// <summary>Clear column grouping metadata from the target range.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Clear { get; set; }

    /// <summary>Keep hidden columns hidden when clearing column grouping metadata.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter KeepHidden { get; set; }

    /// <summary>Places outline summary columns to the left; explicit false places them to the right.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter SummaryLeft { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        int first = ExcelCommandTargets.Column(StartColumn);
        int last = EndColumn == null ? first : ExcelCommandTargets.Column(EndColumn);
        if (last < first) throw new PSArgumentException("The group end must follow its start.");
        if (MyInvocation.BoundParameters.ContainsKey(nameof(SummaryLeft))) sheet.SetOutlineSummary(summaryRight: !SummaryLeft);
        if (Clear) sheet.ClearColumnGroup(first, last, unhide: !KeepHidden);
        else sheet.GroupColumns(first, last, (byte)OutlineLevel, Collapsed, Hidden);
        if (PassThru) WriteObject(sheet);
    }

}
