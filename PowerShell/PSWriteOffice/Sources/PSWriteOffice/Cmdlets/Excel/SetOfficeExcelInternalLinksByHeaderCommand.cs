using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Converts cells under a header into internal workbook links.</summary>
/// <example>
///   <summary>Link the Sheet column in the used range.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Summary' { Set-OfficeExcelInternalLinksByHeader -Header 'Sheet' }</code>
///   <para>Uses the used range header row to find the Sheet column and converts its values into internal links.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelInternalLinksByHeader", DefaultParameterSetName = ParameterSetContextUsedRange)]
[Alias("ExcelInternalLinksByHeader")]
[OutputType(typeof(ExcelSheet))]
public sealed class SetOfficeExcelInternalLinksByHeaderCommand : PSCmdlet
{
    private const string ParameterSetContextUsedRange = "ContextUsedRange";
    private const string ParameterSetDocumentUsedRange = "DocumentUsedRange";
    private const string ParameterSetContextTable = "ContextTable";
    private const string ParameterSetDocumentTable = "DocumentTable";
    private const string ParameterSetContextRange = "ContextRange";
    private const string ParameterSetDocumentRange = "DocumentRange";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentUsedRange)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentTable)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentRange)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentUsedRange)]
    [Parameter(ParameterSetName = ParameterSetDocumentTable)]
    [Parameter(ParameterSetName = ParameterSetDocumentRange)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentUsedRange)]
    [Parameter(ParameterSetName = ParameterSetDocumentTable)]
    [Parameter(ParameterSetName = ParameterSetDocumentRange)]
    public int? SheetIndex { get; set; }

    /// <summary>Header text to locate.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Header { get; set; } = string.Empty;

    /// <summary>Restrict linking to a named table.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetContextTable)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentTable)]
    public string TableName { get; set; } = string.Empty;

    /// <summary>Restrict linking to a specific A1 range whose first row contains headers.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetContextRange)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentRange)]
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
        var toSheet = DestinationSheetScript == null ? null : new System.Func<string, string>(text => ExcelTextTransformService.Apply(DestinationSheetScript, text));
        var display = DisplayScript == null ? null : new System.Func<string, string>(text => ExcelTextTransformService.Apply(DisplayScript, text));

        if (toSheet == null)
        {
            toSheet = text => text;
        }

        switch (ParameterSetName)
        {
            case ParameterSetContextTable:
            case ParameterSetDocumentTable:
                var tableRange = TryGetContextTableRange(sheet);
                if (tableRange == null || tableRange.Length == 0)
                {
                    sheet.LinkByHeaderToInternalSheetsInTable(TableName, Header, toSheet, TargetAddress, display, !NoStyle.IsPresent);
                }
                else
                {
                    sheet.LinkByHeaderToInternalSheetsInRange(tableRange, Header, toSheet, TargetAddress, display, !NoStyle.IsPresent);
                }

                break;
            case ParameterSetContextRange:
            case ParameterSetDocumentRange:
                sheet.LinkByHeaderToInternalSheetsInRange(Range, Header, toSheet, TargetAddress, display, !NoStyle.IsPresent);
                break;
            default:
                sheet.LinkByHeaderToInternalSheets(Header, destinationSheetForCellText: toSheet, targetA1: TargetAddress, display: display, styled: !NoStyle.IsPresent);
                break;
        }

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocumentUsedRange ||
            ParameterSetName == ParameterSetDocumentTable ||
            ParameterSetName == ParameterSetDocumentRange)
        {
            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        return ExcelDslContext.Require(this).RequireSheet();
    }

    private string? TryGetContextTableRange(ExcelSheet sheet)
    {
        if (ParameterSetName != ParameterSetContextTable)
        {
            return null;
        }

        try
        {
            var context = ExcelDslContext.Require(this);
            return context.TryGetTableRange(sheet, TableName, out var range) ? range : null;
        }
        catch (System.InvalidOperationException)
        {
            return null;
        }
    }
}
