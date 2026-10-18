using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets the page orientation on a worksheet.</summary>
/// <example>
///   <summary>Switch to landscape orientation.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelOrientation -Orientation Landscape }</code>
///   <para>Configures the sheet to print in landscape.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelOrientation", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelOrientation")]
public sealed class SetOfficeExcelOrientationCommand : PSCmdlet
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

    /// <summary>Orientation to apply.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public OfficePageOrientation Orientation { get; set; }

    /// <summary>Emit the worksheet after applying the orientation.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        sheet.SetOrientation(Orientation);

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
