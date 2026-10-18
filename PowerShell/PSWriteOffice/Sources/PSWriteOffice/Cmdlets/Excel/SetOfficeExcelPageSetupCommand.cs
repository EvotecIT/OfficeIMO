using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures page setup options on a worksheet.</summary>
/// <example>
///   <summary>Fit to one page wide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelPageSetup -FitToWidth 1 -FitToHeight 0 }</code>
///   <para>Fits the sheet to one page wide and unlimited height.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelPageSetup", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelPageSetup")]
public sealed class SetOfficeExcelPageSetupCommand : PSCmdlet
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

    /// <summary>Number of pages to fit horizontally.</summary>
    [Parameter]
    public uint? FitToWidth { get; set; }

    /// <summary>Number of pages to fit vertically.</summary>
    [Parameter]
    public uint? FitToHeight { get; set; }

    /// <summary>Manual scale percentage (10-400).</summary>
    [Parameter]
    public uint? Scale { get; set; }

    /// <summary>Multi-page print order.</summary>
    [Parameter]
    public ExcelPageOrder? PageOrder { get; set; }

    /// <summary>Emit the worksheet after applying settings.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (FitToWidth == null && FitToHeight == null && Scale == null && PageOrder == null)
        {
            throw new PSArgumentException("Provide FitToWidth, FitToHeight, Scale, or PageOrder.");
        }

        var sheet = ResolveSheet();
        sheet.SetPageSetup(FitToWidth, FitToHeight, Scale, PageOrder);

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
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
