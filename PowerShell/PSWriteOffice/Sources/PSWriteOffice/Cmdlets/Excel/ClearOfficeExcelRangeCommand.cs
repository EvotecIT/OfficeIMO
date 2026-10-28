using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Clears values, formulas, styles, and range metadata from an Excel worksheet range.</summary>
/// <example>
///   <summary>Clear cell contents without removing formatting.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Clear-OfficeExcelRange -Path .\Report.xlsx -Sheet Staging -Range B2:D20 -Contents -Hyperlinks -Confirm:$false
/// Get-OfficeExcelRange -Path .\Report.xlsx -Sheet Staging -Range B2:D20 |
///     Select-Object Address, Value, Formula</code>
///   <para>Removes values and formulas from the selected range and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsCommon.Clear, "OfficeExcelRange", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Medium, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelRangeClear")]
public sealed class ClearOfficeExcelRangeCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name to update. Defaults to the current DSL sheet or the first workbook sheet.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) to update. Defaults to the current DSL sheet or the first workbook sheet.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>A1 range to clear.</summary>
    [Parameter(Mandatory = true)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Clear values and formulas.</summary>
    [Parameter]
    public SwitchParameter Contents { get; set; }

    /// <summary>Clear literal cell values.</summary>
    [Parameter]
    public SwitchParameter Values { get; set; }

    /// <summary>Clear formulas.</summary>
    [Parameter]
    public SwitchParameter Formulas { get; set; }

    /// <summary>Clear cell style indexes.</summary>
    [Parameter]
    [Alias("Formats")]
    public SwitchParameter Styles { get; set; }

    /// <summary>Clear comments in the selected range.</summary>
    [Parameter]
    public SwitchParameter Comments { get; set; }

    /// <summary>Clear hyperlinks that overlap the selected range.</summary>
    [Parameter]
    public SwitchParameter Hyperlinks { get; set; }

    /// <summary>Clear data validation rules that overlap the selected range.</summary>
    [Parameter]
    [Alias("Validation", "Validations")]
    public SwitchParameter DataValidations { get; set; }

    /// <summary>Clear conditional formatting rules that overlap the selected range.</summary>
    [Parameter]
    [Alias("ConditionalFormats")]
    public SwitchParameter ConditionalFormatting { get; set; }

    /// <summary>Clear merged-cell definitions that overlap the selected range.</summary>
    [Parameter]
    public SwitchParameter Merges { get; set; }

    /// <summary>Clear sparklines whose target cells overlap the selected range.</summary>
    [Parameter]
    public SwitchParameter Sparklines { get; set; }

    /// <summary>Clear all supported cell data and range metadata.</summary>
    [Parameter]
    public SwitchParameter All { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = ResolveOptions();
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        var sheet = ExcelWorkbookCommandService.ResolveSheet(this, workbook.Document, ParameterSetName, Sheet, SheetIndex);
        var target = $"{sheet.Name}!{Range}";

        if (ShouldProcess(target, $"Clear Excel range ({options})")) {
            sheet.ClearRange(Range, options);
            workbook.SaveIfOwned();
            WritePassThru(workbook.Document, workbook.OwnsDocument, Path);
        }
    }

    private ExcelClearOptions ResolveOptions() {
        if (All.IsPresent || !AnyOptionSwitchPresent()) {
            return ExcelClearOptions.All;
        }

        var options = ExcelClearOptions.None;
        if (Contents.IsPresent) {
            options |= ExcelClearOptions.Values | ExcelClearOptions.Formulas;
        }

        if (Values.IsPresent) options |= ExcelClearOptions.Values;
        if (Formulas.IsPresent) options |= ExcelClearOptions.Formulas;
        if (Styles.IsPresent) options |= ExcelClearOptions.Styles;
        if (Comments.IsPresent) options |= ExcelClearOptions.Comments;
        if (Hyperlinks.IsPresent) options |= ExcelClearOptions.Hyperlinks;
        if (DataValidations.IsPresent) options |= ExcelClearOptions.DataValidations;
        if (ConditionalFormatting.IsPresent) options |= ExcelClearOptions.ConditionalFormatting;
        if (Merges.IsPresent) options |= ExcelClearOptions.Merges;
        if (Sparklines.IsPresent) options |= ExcelClearOptions.Sparklines;
        return options;
    }

    private bool AnyOptionSwitchPresent() {
        return Contents.IsPresent
            || Values.IsPresent
            || Formulas.IsPresent
            || Styles.IsPresent
            || Comments.IsPresent
            || Hyperlinks.IsPresent
            || DataValidations.IsPresent
            || ConditionalFormatting.IsPresent
            || Merges.IsPresent
            || Sparklines.IsPresent;
    }
}