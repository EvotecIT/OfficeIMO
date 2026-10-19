using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Shows or hides a worksheet.</summary>
/// <example>
///   <summary>Hide the current worksheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelSheetVisibility -Hide }</code>
///   <para>Marks the sheet as hidden.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelSheetVisibility", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelSheetVisibility")]
public sealed class SetOfficeExcelSheetVisibilityCommand : PSCmdlet
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

    /// <summary>Show the worksheet.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Hide the worksheet.</summary>
    [Parameter]
    public SwitchParameter Hide { get; set; }

    /// <summary>Emit the worksheet after applying visibility.</summary>
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
        sheet.SetHidden(Hide.IsPresent);

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
