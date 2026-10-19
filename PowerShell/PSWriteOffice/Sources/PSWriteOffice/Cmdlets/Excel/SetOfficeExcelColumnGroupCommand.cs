using System;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures collapsible Excel outline grouping for worksheet columns.</summary>
/// <example>
///   <summary>Group detail columns and collapse them.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelColumnGroup -StartColumn B -EndColumn D -Collapsed }</code>
///   <para>Applies Excel column outline metadata using OfficeIMO.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelColumnGroup")]
[Alias("ExcelColumnGroup")]
public sealed class SetOfficeExcelColumnGroupCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>First 1-based column in the group.</summary>
    [Parameter(Position = 0)]
    [Alias("Column")]
    public object? StartColumn { get; set; }

    /// <summary>First column letter in the group.</summary>
    [Parameter]
    [Alias("StartColumnLetter", "ColumnName", "ColumnLetter", "Letter")]
    public string? StartColumnName { get; set; }

    /// <summary>Last 1-based column in the group. Defaults to the start column.</summary>
    [Parameter(Position = 1)]
    public object? EndColumn { get; set; }

    /// <summary>Last column letter in the group.</summary>
    [Parameter]
    [Alias("EndColumnLetter")]
    public string? EndColumnName { get; set; }

    /// <summary>Excel outline level from 1 through 7.</summary>
    [Parameter]
    public int OutlineLevel { get; set; } = 1;

    /// <summary>Hide the grouped columns and mark the following summary column as collapsed.</summary>
    [Parameter]
    public SwitchParameter Collapsed { get; set; }

    /// <summary>Hide the grouped columns without marking the group collapsed.</summary>
    [Parameter]
    public SwitchParameter Hidden { get; set; }

    /// <summary>Clear column grouping metadata from the target range.</summary>
    [Parameter]
    public SwitchParameter Clear { get; set; }

    /// <summary>Keep hidden columns hidden when clearing column grouping metadata.</summary>
    [Parameter]
    public SwitchParameter KeepHidden { get; set; }

    /// <summary>Set whether column summary controls appear to the right of grouped columns.</summary>
    [Parameter]
    public bool? SummaryRight { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        var sheet = context.RequireSheet();

        var startColumn = ResolveColumn(StartColumn, StartColumnName, nameof(StartColumn));
        var endColumn = EndColumn != null || !string.IsNullOrWhiteSpace(EndColumnName)
            ? ResolveColumn(EndColumn, EndColumnName, nameof(EndColumn))
            : startColumn;

        if (OutlineLevel < 1 || OutlineLevel > 7) {
            throw new PSArgumentOutOfRangeException(nameof(OutlineLevel), OutlineLevel, "Excel outline level must be between 1 and 7.");
        }

        if (SummaryRight.HasValue) {
            sheet.SetOutlineSummary(summaryRight: SummaryRight.Value);
        }

        if (Clear.IsPresent) {
            sheet.ClearColumnGroup(startColumn, endColumn, unhide: !KeepHidden.IsPresent);
            WritePassThru(sheet);
            return;
        }

        sheet.GroupColumns(startColumn, endColumn, (byte)OutlineLevel, Collapsed.IsPresent, Hidden.IsPresent);
        WritePassThru(sheet);
    }

    private static int ResolveColumn(object? indexOrName, string? name, string parameterName) {
        int? index = null;
        string? columnName = name;

        if (indexOrName is int intValue) {
            index = intValue;
        } else if (indexOrName is long longValue) {
            index = checked((int)longValue);
        } else if (indexOrName != null) {
            var text = Convert.ToString(indexOrName, CultureInfo.InvariantCulture);
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)) {
                index = parsed;
            } else {
                columnName = text;
            }
        }

        try {
            return ExcelHostExtensions.ResolveColumnIndex(index, columnName);
        } catch (Exception exception) when (exception is ArgumentException or FormatException) {
            throw new PSArgumentException(exception.Message, parameterName);
        }
    }
}