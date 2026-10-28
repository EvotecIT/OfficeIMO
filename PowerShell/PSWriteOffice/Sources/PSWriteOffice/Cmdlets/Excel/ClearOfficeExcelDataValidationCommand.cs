using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Clears data validation rules from one or more Excel worksheets.</summary>
/// <example>
///   <summary>Clear validations from a target range.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Clear-OfficeExcelDataValidation -Path .\Report.xlsx -Sheet Data -HeaderName Owner -TableName ServiceHealth -Confirm:$false
/// Get-OfficeExcelDataValidation -Path .\Report.xlsx -Sheet Data |
///     Where-Object Range -like '*Owner*'</code>
///   <para>Removes validation metadata that overlaps the target range and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsCommon.Clear, "OfficeExcelDataValidation", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Medium, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelDataValidationClear")]
public sealed class ClearOfficeExcelDataValidationCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
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

    /// <summary>Worksheet name to update. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) to update. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>Optional A1 range to clear. When omitted, all data validations on the selected sheet are cleared.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Header or table column name used to resolve the range to clear.</summary>
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
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        var shouldSave = false;

        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, workbook.Document, ParameterSetName, Sheet, SheetIndex)) {
            string? targetRange = ExcelTargetRangeResolver.ResolveOptional(sheet, Range, HeaderName, TableName, HeaderRow, IncludeHeader.IsPresent);
            var target = string.IsNullOrWhiteSpace(targetRange) ? sheet.Name : $"{sheet.Name}!{targetRange}";
            if (!ShouldProcess(target, "Clear Excel data validations")) {
                continue;
            }

            shouldSave = true;
            sheet.RemoveDataValidations(targetRange);
        }

        if (shouldSave) {
            workbook.SaveIfOwned();
            WritePassThru(workbook.Document, workbook.OwnsDocument, Path);
        }
    }
}