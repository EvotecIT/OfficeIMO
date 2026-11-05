using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates a report worksheet and lays out report blocks inside it.</summary>
/// <example>
/// <summary>Create a report sheet</summary>
/// <code>New-OfficeExcel -Path .\Reports\Operations.xlsx {
///     ExcelReportSheet 'Summary' {
///         ExcelReportTitle 'Operations summary' 'Week 43, 2026'
///         ExcelReportKpi ([ordered]@{ Uptime = '99.97%'; Incidents = 4; 'Changes shipped' = 37 })
///     }
/// }</code>
/// <para>Adds a Summary report sheet with a title and a KPI row, laid out by the report builder in the workbook theme.</para>
/// </example>
/// <example>
/// <summary>Build a monthly service review</summary>
/// <code>$sites = Import-OfficeCsv -Path .\Data\Site-Availability.csv
/// $legend = @(
///     [pscustomobject]@{ Status = 'Met'; Meaning = 'Availability at or above the 99.9% target' }
///     [pscustomobject]@{ Status = 'Missed'; Meaning = 'Below target; a service credit may apply' }
/// )
/// New-OfficeExcel -Path .\Reports\Service-Review-October.xlsx -Theme Boardroom {
///     ExcelReportSheet 'Review' {
///         ExcelReportTitle 'Service review' 'October 2026'
///         ExcelReportKpi ([ordered]@{ Availability = '99.94%'; 'P1 incidents' = 1; 'Mean time to restore' = '42 min'; 'Changes' = 128 })
///         ExcelReportCallout Warning 'Two sites missed target' 'Gdansk and Porto were below 99.9% after the network maintenance on 14 October.'
///         ExcelReportHeading 'Availability by site'
///         ExcelReportTable $sites 'Sites'
///         ExcelReportChart -Type Column -Category Site -Value Availability -Title 'Availability by site'
///         ExcelReportSpacer
///         ExcelReportLegend 'Status' -Header Status, Meaning -InputObject $legend
///     }
/// }</code>
/// <para>Lays out a complete service review: title, KPIs, a callout, a site table with its chart and a legend, all styled by the theme.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportSheet", DefaultParameterSetName = "Default")]
[Alias("ExcelReportSheet")]
public sealed class AddOfficeExcelReportSheetCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Name of the report worksheet to create.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public string Name { get; set; } = null!;

    /// <summary>Excel child commands executed in the current content scope.</summary>
    [Parameter(ParameterSetName = "Default", Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Fill colour for report section headings.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? HeadingFillColor { get; set; }

    /// <summary>Override the key-cell fill color used by KPI and property blocks.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? KeyFillColor { get; set; }

    /// <summary>Retains native column widths without fitting them to the data.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoAutoFit { get; set; }

    /// <summary>Auto-fit row heights during composer finalization.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AutoFitRows { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ExcelCommandTargets.Document(Document);
        using var ownedContext = ExcelDslContext.Current == null ? ExcelDslContext.Enter(document) : null;
        var context = ExcelDslContext.Require(this);
        if (!ReferenceEquals(context.Document, document)) throw new PSArgumentException("The report document differs from the enclosing workbook.");
        var composer = document.AddSheet(Name).Report(new ExcelReportOptions { SectionHeaderFill = HeadingFillColor, KeyFill = KeyFillColor });
        using (context.Push(composer)) using (context.Push(composer.Sheet)) Content?.InvokeReturnAsIs();
        composer.Finish(autoFitColumns: !NoAutoFit, autoFitRows: AutoFitRows);
        if (PassThru) WriteObject(composer.Sheet);
    }

}
