#pragma warning disable CS1591
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Renames a workbook or sheet-scoped Excel named range.</summary>
/// <example>
///   <summary>Rename a named range and keep the workbook reusable.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$workbook = Get-OfficeExcel -Path .\Report.xlsx
/// $renamed = $workbook | Rename-OfficeExcelNamedRange -Name RevenueRange -NewName Revenue_Current -PassThru
/// Save-OfficeExcel -Document $workbook</code>
///   <para>Renames the defined name through OfficeIMO validation before saving the workbook.</para>
/// </example>
[Cmdlet(VerbsCommon.Rename, "OfficeExcelNamedRange", SupportsShouldProcess = true, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelNamedRangeRename")]
[OutputType(typeof(bool))]
public sealed class RenameOfficeExcelNamedRangeCommand : PSCmdlet
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

    /// <summary>New named range name.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string NewName { get; set; } = string.Empty;

    /// <summary>Use workbook-global scope from inside the DSL.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    public SwitchParameter Global { get; set; }

    /// <summary>Defined-name validation mode.</summary>
    [Parameter]
    public ExcelDefinedNameValidationMode ValidationMode { get; set; } = ExcelDefinedNameValidationMode.Sanitize;
    /// <summary>Emit a result object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Save the workbook immediately after renaming the name.</summary>
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

        if (ShouldProcess(Name, "Rename Excel named range"))
        {
            var renamed = document.RenameNamedRange(Name, NewName, scope, ValidationMode, save: false);
            if (renamed && Save.IsPresent)
            {
                document.Save();
            }
            if (PassThru.IsPresent)
            {
                WriteObject(renamed);
            }
        }
    }
}
#pragma warning restore CS1591
