using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets worksheet view options such as gridlines, direction, zoom, and view mode.</summary>
/// <example>
///   <summary>Apply common worksheet view options.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelWorksheetView -HideGridlines -ZoomScale 125 -View PageLayout }</code>
///   <para>Hides gridlines, sets zoom, and switches the sheet view.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelWorksheetView", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelSheetView")]
public sealed class SetOfficeExcelWorksheetViewCommand : PSCmdlet
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

    /// <summary>Show worksheet gridlines.</summary>
    [Parameter]
    public SwitchParameter ShowGridlines { get; set; }

    /// <summary>Hide worksheet gridlines.</summary>
    [Parameter]
    public SwitchParameter HideGridlines { get; set; }

    /// <summary>Show worksheet right-to-left.</summary>
    [Parameter]
    public SwitchParameter RightToLeft { get; set; }

    /// <summary>Show worksheet left-to-right.</summary>
    [Parameter]
    public SwitchParameter LeftToRight { get; set; }

    /// <summary>Active worksheet zoom percentage. Excel supports values from 10 to 400.</summary>
    [Parameter]
    public uint? ZoomScale { get; set; }

    /// <summary>Normal-view worksheet zoom percentage. Excel supports values from 10 to 400.</summary>
    [Parameter]
    public uint? ZoomScaleNormal { get; set; }

    /// <summary>Worksheet view mode.</summary>
    [Parameter]
    public ExcelWorksheetViewKind? View { get; set; }

    /// <summary>Emit the worksheet after applying view options.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ShowGridlines.IsPresent && HideGridlines.IsPresent)
        {
            throw new PSArgumentException("Specify either -ShowGridlines or -HideGridlines, not both.");
        }

        if (RightToLeft.IsPresent && LeftToRight.IsPresent)
        {
            throw new PSArgumentException("Specify either -RightToLeft or -LeftToRight, not both.");
        }

        if (!HasAnyOption())
        {
            throw new PSArgumentException("Specify at least one worksheet view option.");
        }

        var options = new ExcelWorksheetViewOptions
        {
            ShowGridlines = ShowGridlines.IsPresent ? true : HideGridlines.IsPresent ? false : null,
            RightToLeft = RightToLeft.IsPresent ? true : LeftToRight.IsPresent ? false : null,
            ZoomScale = ZoomScale,
            ZoomScaleNormal = ZoomScaleNormal,
            View = View,
        };

        var sheet = ResolveSheet();
        sheet.SetViewOptions(options);

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private bool HasAnyOption()
    {
        return ShowGridlines.IsPresent ||
            HideGridlines.IsPresent ||
            RightToLeft.IsPresent ||
            LeftToRight.IsPresent ||
            ZoomScale.HasValue ||
            ZoomScaleNormal.HasValue ||
            View.HasValue;
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
