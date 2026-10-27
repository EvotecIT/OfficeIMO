using System.Management.Automation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates a worksheet through the OfficeIMO sheet composer and runs report-block cmdlets inside it.</summary>
/// <example>
///   <summary>Create a composed report sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\report.xlsx {
///     Add-OfficeExcelReportSheet -Name Summary {
///       Add-OfficeExcelReportTitle -Title 'Operational Summary' -Subtitle 'Current view'
///       Add-OfficeExcelReportKpiRow -InputObject @{ Ready = 12; Blocked = 2 }
///     }
///   }</code>
///   <para>Creates a report-oriented worksheet with title and KPI blocks.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportSheet")]
[Alias("ExcelReportSheet")]
public sealed class AddOfficeExcelReportSheetCommand : PSCmdlet
{
    /// <summary>Name of the report worksheet to create.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Report block script to run inside the composer context.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Override the section-header fill color.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? SectionHeaderFillColor { get; set; }

    /// <summary>Override the key-cell fill color used by KPI and property blocks.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? KeyFillColor { get; set; }

    /// <summary>Skip composer auto-fit finalization.</summary>
    [Parameter]
    public SwitchParameter NoAutoFit { get; set; }

    /// <summary>Auto-fit row heights during composer finalization.</summary>
    [Parameter]
    public SwitchParameter AutoFitRows { get; set; }

    /// <summary>Emit the created worksheet.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = ExcelDslContext.Require(this);
        var composer = context.Document.AddWorksheet(Name, ExcelSheetNameValidationMode.Sanitize).Report(BuildTheme());

        using (context.Push(composer))
        using (context.Push(composer.Sheet))
        {
            Content?.InvokeReturnAsIs();
        }

        composer.Finish(autoFitColumns: !NoAutoFit.IsPresent, autoFitRows: AutoFitRows.IsPresent);

        if (PassThru.IsPresent)
        {
            WriteObject(composer.Sheet);
        }
    }

    private ExcelReportOptions BuildTheme()
    {
        if (string.IsNullOrWhiteSpace(SectionHeaderFillColor) && string.IsNullOrWhiteSpace(KeyFillColor))
        {
            return new ExcelReportOptions();
        }

        var theme = new ExcelReportOptions();
        if (!string.IsNullOrWhiteSpace(SectionHeaderFillColor))
        {
            theme.SectionHeaderFill = OfficeIMO.OfficeColor.ParseHex(PSWriteOffice.Services.Text.OfficeColorUtilities.ToExcelColorHex(SectionHeaderFillColor!)!.Substring(2));
        }

        if (!string.IsNullOrWhiteSpace(KeyFillColor))
        {
            theme.KeyFill = OfficeIMO.OfficeColor.ParseHex(PSWriteOffice.Services.Text.OfficeColorUtilities.ToExcelColorHex(KeyFillColor!)!.Substring(2));
        }

        return theme;
    }
}
