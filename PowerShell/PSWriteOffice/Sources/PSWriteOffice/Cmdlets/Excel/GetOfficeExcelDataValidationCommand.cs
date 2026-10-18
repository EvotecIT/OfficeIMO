using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets data validation rules from one or more Excel worksheets.</summary>
/// <example>
///   <summary>Inspect validation rules in a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeExcelDataValidation -Path .\Report.xlsx -Sheet Data |
///     Select-Object -Property SheetName, Range, Type, Operator, Formula1, Formula2</code>
///   <para>Returns validation metadata for workbook audits and maintenance workflows.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelDataValidation", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelDataValidation")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelDataValidationCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to inspect.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to inspect outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name to inspect. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) to inspect. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>Optional A1 range filter.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Header or table column name used to resolve the range filter.</summary>
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

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var path = string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase)
            ? Path
            : null;

        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, workbook.Document, ParameterSetName, Sheet, SheetIndex)) {
            string? targetRange = ExcelTargetRangeResolver.ResolveOptional(sheet, Range, HeaderName, TableName, HeaderRow, IncludeHeader.IsPresent);
            foreach (var validation in sheet.GetDataValidations(targetRange)) {
                WriteObject(ExcelRuleRecordService.CreateDataValidationRecord(validation, sheet.Name, path));
            }
        }
    }
}