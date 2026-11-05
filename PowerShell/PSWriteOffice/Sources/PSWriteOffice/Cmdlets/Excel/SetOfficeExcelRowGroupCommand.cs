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

/// <summary>Groups rows into a collapsible outline, or clears grouping.</summary>
/// <example>
/// <summary>Group detail rows</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Budget-2027.xlsx {
///     ExcelSheet 'Departments' {
///         ExcelRowGroup 3 12
///     }
/// }</code>
/// <para>Groups rows 3 to 12 so they can be collapsed under their summary row.</para>
/// </example>
/// <example>
/// <summary>Collapse each region's detail</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Regional-Sales.xlsx {
///     ExcelSheet 'Detail' {
///         ExcelRowGroup 3 40 -Collapsed -SummaryAbove
///         ExcelRowGroup 42 77 -Collapsed -SummaryAbove
///         ExcelRowGroup 79 120 -Collapsed -SummaryAbove
///     }
/// }</code>
/// <para>Groups and collapses the detail rows of each region with the summary row above them, so the sheet opens showing only the regional totals.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelRowGroup", DefaultParameterSetName = "Default")]
[Alias("ExcelRowGroup")]
public sealed class SetOfficeExcelRowGroupCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>First 1-based row in the group.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public int StartRow { get; set; }

    /// <summary>Last 1-based row in the group. Defaults to <see cref="StartRow"/>.</summary>
    [Parameter(ParameterSetName = "Default", Position = 1)]
    public int? EndRow { get; set; }

    /// <summary>Excel outline level from 1 through 7.</summary>
    [Parameter(ParameterSetName = "Default")]
    [ValidateRange(1, 7)]
    public int OutlineLevel { get; set; } = 1;

    /// <summary>Hide the grouped rows and mark the following summary row as collapsed.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Collapsed { get; set; }

    /// <summary>Hide the grouped rows without marking the group collapsed.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Hidden { get; set; }

    /// <summary>Clear row grouping metadata from the target range.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Clear { get; set; }

    /// <summary>Keep hidden rows hidden when clearing row grouping metadata.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter KeepHidden { get; set; }

    /// <summary>Places outline summary rows above detail rows; explicit false places them below.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter SummaryAbove { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        int first = StartRow;
        int last = EndRow ?? first;
        if (last < first) throw new PSArgumentException("The group end must follow its start.");
        if (MyInvocation.BoundParameters.ContainsKey(nameof(SummaryAbove))) sheet.SetOutlineSummary(summaryBelow: !SummaryAbove);
        if (Clear) sheet.ClearRowGroup(first, last, unhide: !KeepHidden);
        else sheet.GroupRows(first, last, (byte)OutlineLevel, Collapsed, Hidden);
        if (PassThru) WriteObject(sheet);
    }

}
