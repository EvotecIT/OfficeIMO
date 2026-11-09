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

/// <summary>Sets prompt and error messages on existing validation rules.</summary>
/// <example>
/// <summary>Add an input prompt</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
///     ExcelSheet 'Tickets' {
///         Set-OfficeExcelValidation -Header Priority -Table Tickets -PromptTitle 'Priority' -Prompt 'P1 is a production outage; P4 is a question.' -ShowInputMessage
///     }
/// }</code>
/// <para>Shows a hint when someone selects a cell in the Priority column.</para>
/// </example>
/// <example>
/// <summary>Reword error messages across a form</summary>
/// <code>Open-OfficeExcel -Path .\Forms\Equipment-Requests.xlsx {
///     ExcelSheet 'Requests' {
///         Set-OfficeExcelValidation -Range E2:E500 -ErrorTitle 'Cost centre' -ErrorMessage 'Pick a cost centre from the list. Ask Finance if yours is missing.' -ShowErrorMessage
///     }
/// }</code>
/// <para>Replaces the generic error text of the cost-centre rule in an existing form with guidance for the requester.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelValidation", DefaultParameterSetName = "Range")]
public sealed class SetOfficeExcelValidationCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter(ParameterSetName = "Range", Position = 0)]
    public string? Range { get; set; }

    /// <summary>Input prompt title. Omit or pass null to clear the title.</summary>
    [Parameter()]
    public string? PromptTitle { get; set; }

    /// <summary>Input prompt text. Omit or pass null to clear the prompt.</summary>
    [Parameter()]
    public string? Prompt { get; set; }

    /// <summary>Error title. Omit or pass null to clear the title.</summary>
    [Parameter()]
    public string? ErrorTitle { get; set; }

    /// <summary>Error message text. Omit or pass null to clear the message.</summary>
    [Parameter()]
    public string? ErrorMessage { get; set; }

    /// <summary>Forces Excel to show the input prompt.</summary>
    [Parameter()]
    public SwitchParameter ShowInputMessage { get; set; }

    /// <summary>Forces Excel to show the validation error.</summary>
    [Parameter()]
    public SwitchParameter ShowErrorMessage { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Header text identifying the target column.</summary>
    [Parameter(ParameterSetName = "Header", Mandatory = true)]
    public string Header { get; set; } = null!;

    /// <summary>Native Excel table used to resolve the source range or header.</summary>
    [Parameter(ParameterSetName = "Header")]
    public string? Table { get; set; }

    /// <summary>One-based row containing the column headers.</summary>
    [Parameter(ParameterSetName = "Header")]
    public int? HeaderRow { get; set; }

    /// <summary>Includes the header cell in the resolved column range.</summary>
    [Parameter(ParameterSetName = "Header")]
    public SwitchParameter IncludeHeader { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!HasMessageOption()) throw new PSArgumentException("Supply a prompt, error or display option.");
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        string targetRange = ExcelTargetRangeResolver.ResolveOptional(sheet, Range, Header, Table, HeaderRow ?? 0, IncludeHeader) ?? sheet.UsedRange.Address;
        sheet.UpdateDataValidationMessages(targetRange, options => {
            bool hadPrompt = HasMessageText(options.PromptTitle, options.Prompt);
            bool hadError = HasMessageText(options.ErrorTitle, options.Error);
            if (MyInvocation.BoundParameters.ContainsKey(nameof(PromptTitle))) options.PromptTitle = PromptTitle;
            if (MyInvocation.BoundParameters.ContainsKey(nameof(Prompt))) options.Prompt = Prompt;
            if (MyInvocation.BoundParameters.ContainsKey(nameof(ErrorTitle))) options.ErrorTitle = ErrorTitle;
            if (MyInvocation.BoundParameters.ContainsKey(nameof(ErrorMessage))) options.Error = ErrorMessage;
            if (MyInvocation.BoundParameters.ContainsKey(nameof(ShowInputMessage))) options.ShowInputMessage = ShowInputMessage.IsPresent;
            else if (!hadPrompt) options.ShowInputMessage = HasMessageText(options.PromptTitle, options.Prompt);
            if (MyInvocation.BoundParameters.ContainsKey(nameof(ShowErrorMessage))) options.ShowErrorMessage = ShowErrorMessage.IsPresent;
            else if (!hadError) options.ShowErrorMessage = HasMessageText(options.ErrorTitle, options.Error);
        });

        if (PassThru) foreach (var rule in sheet.GetDataValidations(targetRange)) WriteObject(rule);
    }
    private bool HasMessageOption() {
        return MyInvocation.BoundParameters.ContainsKey(nameof(PromptTitle))
            || MyInvocation.BoundParameters.ContainsKey(nameof(Prompt))
            || MyInvocation.BoundParameters.ContainsKey(nameof(ErrorTitle))
            || MyInvocation.BoundParameters.ContainsKey(nameof(ErrorMessage))
            || MyInvocation.BoundParameters.ContainsKey(nameof(ShowInputMessage))
            || MyInvocation.BoundParameters.ContainsKey(nameof(ShowErrorMessage));
    }

    private static bool HasMessageText(string? title, string? message)
        => !string.IsNullOrEmpty(title) || !string.IsNullOrEmpty(message);

    private static IReadOnlyList<ExcelDataValidationInfo> GetDataValidations(ExcelSheet sheet, string targetRange)
        => sheet.GetDataValidations(targetRange);

}
