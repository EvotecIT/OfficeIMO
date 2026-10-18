using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets page margins on a worksheet.</summary>
/// <example>
///   <summary>Apply a preset margin set.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelMargins -Preset Narrow }</code>
///   <para>Applies the Narrow margin preset.</para>
/// </example>
/// <example>
///   <summary>Apply custom margins.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelMargins -Left 0.5 -Right 0.5 -Top 0.75 -Bottom 0.75 }</code>
///   <para>Sets custom margins in inches.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelMargins", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelMargins")]
public sealed class SetOfficeExcelMarginsCommand : PSCmdlet
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

    /// <summary>Margin preset to apply.</summary>
    [Parameter]
    public ExcelMarginPreset? Preset { get; set; }

    /// <summary>Left margin in inches.</summary>
    [Parameter]
    public double? Left { get; set; }

    /// <summary>Right margin in inches.</summary>
    [Parameter]
    public double? Right { get; set; }

    /// <summary>Top margin in inches.</summary>
    [Parameter]
    public double? Top { get; set; }

    /// <summary>Bottom margin in inches.</summary>
    [Parameter]
    public double? Bottom { get; set; }

    /// <summary>Header margin in inches.</summary>
    [Parameter]
    public double? Header { get; set; }

    /// <summary>Footer margin in inches.</summary>
    [Parameter]
    public double? Footer { get; set; }

    /// <summary>Emit the worksheet after applying margins.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();

        if (Preset.HasValue)
        {
            if (Left.HasValue || Right.HasValue || Top.HasValue || Bottom.HasValue || Header.HasValue || Footer.HasValue)
            {
                throw new PSArgumentException("Provide either Preset or custom margins, not both.");
            }

            sheet.SetMarginsPreset(Preset.Value);
        }
        else
        {
            if (!Left.HasValue || !Right.HasValue || !Top.HasValue || !Bottom.HasValue)
            {
                throw new PSArgumentException("Provide Left, Right, Top, and Bottom margins or use a Preset.");
            }

            sheet.SetMargins(
                left: Left.Value,
                right: Right.Value,
                top: Top.Value,
                bottom: Bottom.Value,
                header: Header ?? 0.3,
                footer: Footer ?? 0.3);
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
