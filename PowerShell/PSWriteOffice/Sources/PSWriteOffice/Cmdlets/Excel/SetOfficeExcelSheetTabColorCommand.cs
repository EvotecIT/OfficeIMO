using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets or clears the worksheet tab color.</summary>
/// <example>
///   <summary>Set a worksheet tab color.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelSheetTabColor -Color '#336699' }</code>
///   <para>Applies the color to the active worksheet tab.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelSheetTabColor", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelSheetTabColor")]
public sealed class SetOfficeExcelSheetTabColorCommand : PSCmdlet
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

    /// <summary>Named color, RGB hex, or RGBA hex value.</summary>
    [Parameter(Position = 0)]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Clear the worksheet tab color.</summary>
    [Parameter]
    public SwitchParameter Clear { get; set; }

    /// <summary>Emit the worksheet after applying the tab color.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Clear.IsPresent && !string.IsNullOrWhiteSpace(Color))
        {
            throw new PSArgumentException("Specify either -Color or -Clear, not both.");
        }

        if (!Clear.IsPresent && string.IsNullOrWhiteSpace(Color))
        {
            throw new PSArgumentException("Specify -Color or -Clear.");
        }

        var sheet = ResolveSheet();
        if (Clear.IsPresent)
        {
            sheet.ClearTabColor();
        }
        else
        {
            sheet.SetTabColor(Color!);
        }

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
