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

/// <summary>Adds slicer or timeline binding metadata for a pivot table or table.</summary>
/// <example>
/// <summary>Add a slicer for a pivot field</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Sales-2026.xlsx {
///     ExcelSlicer -Name RegionSlicer -SourceName Region -PivotTable SalesPivot
/// }</code>
/// <para>Adds slicer metadata for the Region field of the SalesPivot pivot table.</para>
/// </example>
/// <example>
/// <summary>Add a timeline for order dates</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Sales-2026.xlsx {
///     ExcelSlicer -Name OrderDateTimeline -SourceName OrderDate -PivotTable SalesPivot -Timeline
/// }</code>
/// <para>Adds timeline metadata for the OrderDate field so readers can filter the pivot table by period in Excel.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelSlicer", DefaultParameterSetName = "Default")]
[Alias("ExcelSlicer")]
public sealed class AddOfficeExcelSlicerCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Slicer cache name.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true)]
    public string Name { get; set; } = null!;

    /// <summary>Source field, table column, or pivot field name.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? SourceName { get; set; }

    /// <summary>Native pivot table whose data or entire range is targeted.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? PivotTable { get; set; }

    /// <summary>Caller-supplied slicer cache XML. When provided, OfficeIMO preserves it within its pivot-interaction metadata envelope.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Xml { get; set; }

    /// <summary>Creates timeline metadata instead of a standard slicer.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Timeline { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ExcelCommandTargets.Document(Document);
        object result = Timeline ? document.AddWorkbookTimelineCache(new ExcelTimelineCacheOptions { Name = Name, SourceName = SourceName, PivotTableName = PivotTable, Xml = Xml }) : document.AddWorkbookSlicerCache(new ExcelSlicerCacheOptions { Name = Name, SourceName = SourceName, PivotTableName = PivotTable, Xml = Xml });
        if (PassThru) WriteObject(result);
    }

}
