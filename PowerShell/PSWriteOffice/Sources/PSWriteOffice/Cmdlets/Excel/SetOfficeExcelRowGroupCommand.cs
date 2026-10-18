using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures collapsible Excel outline grouping for worksheet rows.</summary>
/// <example>
///   <summary>Group detail rows and collapse them.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelRowGroup -StartRow 2 -EndRow 20 -Collapsed }</code>
///   <para>Applies Excel row outline metadata using OfficeIMO.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelRowGroup")]
[Alias("ExcelRowGroup")]
public sealed class SetOfficeExcelRowGroupCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>First 1-based row in the group.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public int StartRow { get; set; }

    /// <summary>Last 1-based row in the group. Defaults to <see cref="StartRow"/>.</summary>
    [Parameter(Position = 1)]
    public int? EndRow { get; set; }

    /// <summary>Excel outline level from 1 through 7.</summary>
    [Parameter]
    public int OutlineLevel { get; set; } = 1;

    /// <summary>Hide the grouped rows and mark the following summary row as collapsed.</summary>
    [Parameter]
    public SwitchParameter Collapsed { get; set; }

    /// <summary>Hide the grouped rows without marking the group collapsed.</summary>
    [Parameter]
    public SwitchParameter Hidden { get; set; }

    /// <summary>Clear row grouping metadata from the target range.</summary>
    [Parameter]
    public SwitchParameter Clear { get; set; }

    /// <summary>Keep hidden rows hidden when clearing row grouping metadata.</summary>
    [Parameter]
    public SwitchParameter KeepHidden { get; set; }

    /// <summary>Set whether row summary controls appear below grouped rows.</summary>
    [Parameter]
    public bool? SummaryBelow { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        var sheet = context.RequireSheet();
        var lastRow = EndRow ?? StartRow;

        if (OutlineLevel < 1 || OutlineLevel > 7) {
            throw new PSArgumentOutOfRangeException(nameof(OutlineLevel), OutlineLevel, "Excel outline level must be between 1 and 7.");
        }

        if (SummaryBelow.HasValue) {
            sheet.SetOutlineSummary(summaryBelow: SummaryBelow.Value);
        }

        if (Clear.IsPresent) {
            sheet.ClearRowGroup(StartRow, lastRow, unhide: !KeepHidden.IsPresent);
            WritePassThru(sheet);
            return;
        }

        sheet.GroupRows(StartRow, lastRow, (byte)OutlineLevel, Collapsed.IsPresent, Hidden.IsPresent);
        WritePassThru(sheet);
    }
}