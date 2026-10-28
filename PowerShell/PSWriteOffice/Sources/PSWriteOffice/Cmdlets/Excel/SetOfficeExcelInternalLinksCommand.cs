using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Converts cells in a range into internal workbook links.</summary>
/// <example>
///   <summary>Link a summary range to same-named sheets.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Summary' { Set-OfficeExcelInternalLinks -Range 'A2:A10' }</code>
///   <para>Turns each non-empty cell in A2:A10 into an internal link to the sheet with the same name.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelInternalLinks", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelInternalLinks")]
[OutputType(typeof(ExcelSheet))]
public sealed class SetOfficeExcelInternalLinksCommand : PSCmdlet
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

    /// <summary>A1 range containing values to convert into internal links.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Optional mapping from cell text to destination sheet name.</summary>
    [Parameter]
    public ScriptBlock? DestinationSheetScript { get; set; }

    /// <summary>Optional mapping from cell text to display text.</summary>
    [Parameter]
    public ScriptBlock? DisplayScript { get; set; }

    /// <summary>Destination cell on the target sheet.</summary>
    [Parameter]
    public string TargetAddress { get; set; } = "A1";

    /// <summary>Skip hyperlink styling (blue + underline).</summary>
    [Parameter]
    public SwitchParameter NoStyle { get; set; }

    /// <summary>Emit the worksheet after creating links.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        sheet.LinkCellsToInternalSheets(
            Range,
            text => ExcelTextTransformService.Apply(DestinationSheetScript, text),
            targetA1: TargetAddress,
            display: DisplayScript == null ? null : text => ExcelTextTransformService.Apply(DisplayScript, text),
            styled: !NoStyle.IsPresent);

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        return ExcelDslContext.Require(this).RequireSheet();
    }
}
