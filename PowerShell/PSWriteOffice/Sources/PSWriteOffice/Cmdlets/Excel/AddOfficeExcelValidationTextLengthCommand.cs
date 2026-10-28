using System.Management.Automation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a text-length data validation rule to a worksheet range.</summary>
/// <example>
///   <summary>Restrict values to 1-50 characters.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelValidationTextLength -Range 'E2:E20' -Operator Between -Formula1 1 -Formula2 50 }</code>
///   <para>Ensures text length in E2:E20 is between 1 and 50 characters.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelValidationTextLength", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelValidationTextLength")]
public sealed class AddOfficeExcelValidationTextLengthCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>Target range in A1 notation.</summary>
    [Parameter(Position = 0)]
    public string? Range { get; set; }

    /// <summary>Header or table column name used to resolve the target range.</summary>
    [Parameter]
    [Alias("ColumnName")]
    public string? HeaderName { get; set; }

    /// <summary>Optional table name for header-based range resolution.</summary>
    [Parameter]
    public string? TableName { get; set; }

    /// <summary>Worksheet header row used when resolving HeaderName without a table. Use 0 for the first row of the used range.</summary>
    [Parameter]
    public int HeaderRow { get; set; }

    /// <summary>Include the header cell in the resolved range.</summary>
    [Parameter]
    public SwitchParameter IncludeHeader { get; set; }

    /// <summary>Validation operator.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public ExcelDataValidationOperator Operator { get; set; }

    /// <summary>Primary length bound.</summary>
    [Parameter(Mandatory = true, Position = 2)]
    public int Formula1 { get; set; }

    /// <summary>Optional secondary length bound.</summary>
    [Parameter]
    public int? Formula2 { get; set; }

    /// <summary>Allow blank values.</summary>
    [Parameter]
    public bool AllowBlank { get; set; } = true;

    /// <summary>Error title shown to the user.</summary>
    [Parameter]
    public string? ErrorTitle { get; set; }

    /// <summary>Error message shown to the user.</summary>
    [Parameter]
    public string? ErrorMessage { get; set; }

    /// <summary>Emit the range string after applying validation.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        string targetRange = ExcelTargetRangeResolver.Resolve(sheet, Range, HeaderName, TableName, HeaderRow, IncludeHeader.IsPresent);
        sheet.ValidationTextLength(targetRange, Operator, Formula1, Formula2, AllowBlank, ErrorTitle, ErrorMessage);

        if (PassThru.IsPresent)
        {
            WriteObject(targetRange);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            if (Document == null)
            {
                throw new PSArgumentException("Provide an Excel document.");
            }

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }
}
