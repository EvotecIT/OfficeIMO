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

/// <summary>Adds a data validation rule to a range or a header column.</summary>
/// <example>
/// <summary>Restrict a column to a list</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
///     ExcelSheet 'Tickets' {
///         ExcelValidation -Range C2:C500 -Type List -Values 'Open', 'In progress', 'On hold', 'Closed'
///     }
/// }</code>
/// <para>Adds a drop-down that allows only the listed statuses in C2:C500.</para>
/// </example>
/// <example>
/// <summary>Validate quantities in a request form</summary>
/// <code>Open-OfficeExcel -Path .\Forms\Equipment-Requests.xlsx {
///     ExcelSheet 'Requests' {
///         ExcelValidation -Header Quantity -Table Requests -Type WholeNumber -Operator Between -Value 1 -SecondValue 500 -NoBlank -ErrorTitle 'Quantity' -ErrorMessage 'Enter a whole number from 1 to 500.'
///     }
/// }</code>
/// <para>Allows whole numbers from 1 to 500 in the Quantity column of the Requests table and explains the rule when a wrong value is typed.</para>
/// </example>
/// <example>
/// <summary>Validate with a formula</summary>
/// <code>Open-OfficeExcel -Path .\Projects\Plan.xlsx {
///     ExcelSheet 'Tasks' {
///         ExcelValidation -Range D2:D300 -Type Formula -Formula '=D2&gt;=C2' -ErrorMessage 'The end date must be on or after the start date.'
///     }
/// }</code>
/// <para>Rejects end dates that come before the start date in the same row.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelValidation", DefaultParameterSetName = "ListRange")]
[Alias("ExcelValidation")]
public sealed class AddOfficeExcelValidationCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter(ParameterSetName = "ListRange", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "CompareRange", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "FormulaRange", Mandatory = true, Position = 0)]
    public string Range { get; set; } = null!;

    /// <summary>Allowed values for the dropdown list.</summary>
    [Parameter(ParameterSetName = "ListRange", Mandatory = true, Position = 2)]
    [Parameter(ParameterSetName = "ListHeader", Mandatory = true, Position = 2)]
    public string[] Values { get; set; } = Array.Empty<string>();

    /// <summary>Shared chart family. Bubble requires the separate native bubble-data contract.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public ExcelValidationType Type { get; set; }

    /// <summary>Rejects blank input in the validation rule.</summary>
    [Parameter()]
    public SwitchParameter NoBlank { get; set; }

    /// <summary>Input prompt title. Omit or pass null to clear the title.</summary>
    [Parameter()]
    public string? PromptTitle { get; set; }

    /// <summary>Input prompt text. Omit or pass null to clear the prompt.</summary>
    [Parameter()]
    public string? Prompt { get; set; }

    /// <summary>Error title shown to the user.</summary>
    [Parameter()]
    public string? ErrorTitle { get; set; }

    /// <summary>Error message shown to the user.</summary>
    [Parameter()]
    public string? ErrorMessage { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Header text identifying the target column.</summary>
    [Parameter(ParameterSetName = "ListHeader", Mandatory = true)]
    [Parameter(ParameterSetName = "CompareHeader", Mandatory = true)]
    [Parameter(ParameterSetName = "FormulaHeader", Mandatory = true)]
    public string Header { get; set; } = null!;

    /// <summary>Native Excel table used to resolve the source range or header.</summary>
    [Parameter(ParameterSetName = "ListHeader")]
    [Parameter(ParameterSetName = "CompareHeader")]
    [Parameter(ParameterSetName = "FormulaHeader")]
    public string? Table { get; set; }

    /// <summary>One-based row containing the column headers.</summary>
    [Parameter(ParameterSetName = "ListHeader")]
    [Parameter(ParameterSetName = "CompareHeader")]
    [Parameter(ParameterSetName = "FormulaHeader")]
    public int? HeaderRow { get; set; }

    /// <summary>Includes the header cell in the resolved column range.</summary>
    [Parameter(ParameterSetName = "ListHeader")]
    [Parameter(ParameterSetName = "CompareHeader")]
    [Parameter(ParameterSetName = "FormulaHeader")]
    public SwitchParameter IncludeHeader { get; set; }

    /// <summary>Validation operator.</summary>
    [Parameter(ParameterSetName = "CompareRange", Mandatory = true)]
    [Parameter(ParameterSetName = "CompareHeader", Mandatory = true)]
    public ExcelDataValidationOperator Operator { get; set; }

    /// <summary>Value or values consumed by the selected operation.</summary>
    [Parameter(ParameterSetName = "CompareRange", Mandatory = true, Position = 2)]
    [Parameter(ParameterSetName = "CompareHeader", Mandatory = true, Position = 2)]
    public object Value { get; set; } = null!;

    /// <summary>Upper bound or second formula used by between comparisons.</summary>
    [Parameter(ParameterSetName = "CompareRange")]
    [Parameter(ParameterSetName = "CompareHeader")]
    public object? SecondValue { get; set; }

    /// <summary>Validation formula.</summary>
    [Parameter(ParameterSetName = "FormulaRange", Mandatory = true, Position = 2)]
    [Parameter(ParameterSetName = "FormulaHeader", Mandatory = true, Position = 2)]
    public string Formula { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        string target = ExcelTargetRangeResolver.Resolve(sheet, Range, Header, Table, HeaderRow ?? 0, IncludeHeader);
        var builder = sheet.Range(target).Validation();
        bool blank = !NoBlank;
        if (ParameterSetName.StartsWith("List", StringComparison.Ordinal)) {
            if (Type != ExcelValidationType.List) throw new PSArgumentException("-Values requires -Type List.");
            builder.List(Values, blank);
        } else if (ParameterSetName.StartsWith("Formula", StringComparison.Ordinal)) {
            if (Type != ExcelValidationType.Formula) throw new PSArgumentException("-Formula requires -Type Formula.");
            builder.CustomFormula(Formula!, blank, ErrorTitle, ErrorMessage);
        } else {
            object first = OfficeArgumentValue.Unwrap(Value)!;
            object? second = OfficeArgumentValue.Unwrap(SecondValue);
            switch (Type) {
                case ExcelValidationType.WholeNumber: builder.WholeNumber(Operator, Convert.ToInt32(first, CultureInfo.InvariantCulture), second == null ? null : Convert.ToInt32(second, CultureInfo.InvariantCulture), blank, ErrorTitle, ErrorMessage); break;
                case ExcelValidationType.Decimal: builder.Decimal(Operator, Convert.ToDouble(first, CultureInfo.InvariantCulture), second == null ? null : Convert.ToDouble(second, CultureInfo.InvariantCulture), blank, ErrorTitle, ErrorMessage); break;
                case ExcelValidationType.Date: builder.Date(Operator, Convert.ToDateTime(first, CultureInfo.CurrentCulture), second == null ? null : Convert.ToDateTime(second, CultureInfo.CurrentCulture), blank, ErrorTitle, ErrorMessage); break;
                case ExcelValidationType.Time: builder.Time(Operator, TimeValue(first), second == null ? null : TimeValue(second), blank, ErrorTitle, ErrorMessage); break;
                case ExcelValidationType.TextLength: builder.TextLength(Operator, Convert.ToInt32(first, CultureInfo.InvariantCulture), second == null ? null : Convert.ToInt32(second, CultureInfo.InvariantCulture), blank, ErrorTitle, ErrorMessage); break;
                default: throw new PSArgumentException("Comparison sets require WholeNumber, Decimal, Date, Time or TextLength.");
            }
        }
        if (PromptTitle != null || Prompt != null || ErrorTitle != null || ErrorMessage != null)
            sheet.UpdateDataValidationMessages(target, options => {
                options.PromptTitle = PromptTitle; options.Prompt = Prompt;
                options.ErrorTitle = ErrorTitle; options.Error = ErrorMessage;
                options.ShowInputMessage = PromptTitle != null || Prompt != null;
                options.ShowErrorMessage = ErrorTitle != null || ErrorMessage != null;
            });
        if (PassThru) foreach (var rule in sheet.GetDataValidations(target)) WriteObject(rule);
    }
    private static TimeSpan TimeValue(object value) => value is TimeSpan span ? span : value is DateTime date ? date.TimeOfDay : TimeSpan.Parse(Convert.ToString(value, CultureInfo.InvariantCulture)!, CultureInfo.InvariantCulture);

}
