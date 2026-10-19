using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Shows or hides worksheet gridlines.</summary>
/// <example>
///   <summary>Hide gridlines.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelGridlines -Hide }</code>
///   <para>Turns off gridlines for the sheet.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelGridlines", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelGridlines")]
public sealed class SetOfficeExcelGridlinesCommand : PSCmdlet
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

    /// <summary>Show gridlines.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Hide gridlines.</summary>
    [Parameter]
    public SwitchParameter Hide { get; set; }

    /// <summary>Emit the worksheet after applying gridline settings.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Show.IsPresent && Hide.IsPresent)
        {
            throw new PSArgumentException("Specify either -Show or -Hide, not both.");
        }

        if (!Show.IsPresent && !Hide.IsPresent)
        {
            throw new PSArgumentException("Specify -Show or -Hide.");
        }

        var sheet = ResolveSheet();
        sheet.SetGridlinesVisible(Show.IsPresent);

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
