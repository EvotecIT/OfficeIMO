#pragma warning disable CS1591
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Removes a workbook or sheet-scoped Excel named range.</summary>
/// <example>
///   <summary>Remove a stale sheet-scoped named range from a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$workbook = Get-OfficeExcel -Path .\Report.xlsx
/// $removed = $workbook | Remove-OfficeExcelNamedRange -Sheet Data -Name OldCriteria -PassThru
/// Save-OfficeExcel -Document $workbook</code>
///   <para>Uses the thin PowerShell surface over OfficeIMO named-range removal and saves the updated workbook.</para>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficeExcelNamedRange", SupportsShouldProcess = true, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelNamedRangeRemove")]
[OutputType(typeof(bool))]
public sealed class RemoveOfficeExcelNamedRangeCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name for a sheet-scoped operation.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public string? Sheet { get; set; }

    /// <summary>Zero-based worksheet index for a sheet-scoped operation.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>Named range name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Use workbook-global scope from inside the DSL.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    public SwitchParameter Global { get; set; }

    /// <summary>Emit a result object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Save the workbook immediately after removing the name.</summary>
    [Parameter]
    public SwitchParameter Save { get; set; }

    protected override void ProcessRecord()
    {
        ExcelDocument document;
        ExcelSheet? scope = null;
        if (ParameterSetName == ParameterSetDocument)
        {
            document = Document;
            scope = ExcelSheetResolver.ResolveOptional(document, Sheet, SheetIndex);
        }
        else
        {
            var context = ExcelDslContext.Require(this);
            document = context.Document;
            if (!Global.IsPresent)
            {
                scope = context.CurrentSheet;
            }
        }

        if (ShouldProcess(Name, "Remove Excel named range"))
        {
            var removed = document.RemoveNamedRange(Name, scope, save: false);
            if (removed && Save.IsPresent)
            {
                document.Save();
            }
            if (PassThru.IsPresent)
            {
                WriteObject(removed);
            }
        }
    }
}
#pragma warning restore CS1591
