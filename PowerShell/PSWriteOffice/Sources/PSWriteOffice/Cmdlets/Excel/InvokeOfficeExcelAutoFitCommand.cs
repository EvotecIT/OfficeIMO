using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Automatically fits Excel row heights and/or column widths.</summary>
/// <para>Targets the current worksheet context or a worksheet on a supplied document.</para>
/// <example>
///   <summary>Auto-fit columns in the current DSL sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Invoke-OfficeExcelAutoFit -Columns }</code>
///   <para>Adjusts column widths for the active sheet.</para>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeExcelAutoFit", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelAutoFit")]
public sealed class InvokeOfficeExcelAutoFitCommand : PSCmdlet
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

    /// <summary>Auto-fit all columns.</summary>
    [Parameter]
    public SwitchParameter Columns { get; set; }

    /// <summary>Auto-fit all rows.</summary>
    [Parameter]
    public SwitchParameter Rows { get; set; }

    /// <summary>Auto-fit both rows and columns.</summary>
    [Parameter]
    public SwitchParameter All { get; set; }

    /// <summary>Auto-fit specific column indexes (1-based).</summary>
    [Parameter]
    public int[]? Column { get; set; }

    /// <summary>Auto-fit specific row indexes (1-based).</summary>
    [Parameter]
    public int[]? Row { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();

        var hasColumnList = Column != null && Column.Length > 0;
        var hasRowList = Row != null && Row.Length > 0;

        var autoColumns = All.IsPresent || Columns.IsPresent;
        var autoRows = All.IsPresent || Rows.IsPresent;

        if (!hasColumnList && !hasRowList && !autoColumns && !autoRows)
        {
            autoColumns = true;
        }

        if (hasColumnList)
        {
            sheet.AutoFitColumnsFor(Column!);
        }
        else if (autoColumns)
        {
            sheet.AutoFitColumns();
        }

        if (hasRowList)
        {
            foreach (var rowIndex in Row!.Where(r => r > 0))
            {
                sheet.AutoFitRow(rowIndex);
            }
        }
        else if (autoRows)
        {
            sheet.AutoFitRows();
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
