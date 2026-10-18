using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds OfficeIMO-owned workbook slicer binding metadata.</summary>
/// <example>
///   <summary>Add slicer binding metadata for a pivot field.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$slicer = Add-OfficeExcelSlicer -Path .\Report.xlsx -Name RegionSlicer -SourceName Region -PivotTableName SalesPivot -PassThru
/// Get-OfficeExcelDataModel -Path .\Report.xlsx |
///     Select-Object -ExpandProperty SlicerCacheCount</code>
///   <para>Writes portable OfficeIMO binding metadata. It does not create native Excel slicer caches or UI shapes.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelSlicer", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelSlicer")]
[OutputType(typeof(PSObject))]
public sealed class AddOfficeExcelSlicerCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Slicer cache name.</summary>
    [Parameter(Mandatory = true)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Source field, table column, or pivot field name.</summary>
    [Parameter]
    public string? SourceName { get; set; }

    /// <summary>Pivot table name the slicer is intended to filter.</summary>
    [Parameter]
    public string? PivotTableName { get; set; }

    /// <summary>Caller-supplied slicer cache XML. When provided, OfficeIMO preserves it within its pivot-interaction metadata envelope.</summary>
    [Parameter]
    public string? Xml { get; set; }

    /// <summary>Emit metadata about the added package part.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {
            return;
        }

        ExcelPackagePartInfo part = workbook.Document.AddWorkbookSlicerCache(new ExcelSlicerCacheOptions {
            Name = Name,
            SourceName = SourceName,
            PivotTableName = PivotTableName,
            Xml = Xml
        });

        string contentType = part.ContentType;
        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            var result = new PSObject();
            result.Properties.Add(new PSNoteProperty("Name", Name));
            result.Properties.Add(new PSNoteProperty("Kind", "Slicer"));
            result.Properties.Add(new PSNoteProperty("ContentType", contentType));
            WriteObject(result);
        }
    }
}