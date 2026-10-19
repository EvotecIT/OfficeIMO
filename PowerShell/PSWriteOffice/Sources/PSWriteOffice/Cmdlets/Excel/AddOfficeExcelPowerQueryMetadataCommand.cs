#pragma warning disable CS1591
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds safe Power Query/connection metadata for Excel-compatible applications to own and refresh.</summary>
/// <example>
///   <summary>Add a workbook connection and worksheet query-table metadata.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeExcelPowerQueryMetadata -Path .\Report.xlsx `
///     -Name SalesQuery `
///     -WorksheetName Data `
///     -CommandText 'let Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content] in Source' `
///     -Description 'Sales query metadata authored by automation' `
///     -RefreshOnOpen `
///     -PassThru</code>
///   <para>Writes package metadata only. OfficeIMO does not execute Power Query M; Excel-compatible applications perform refresh when opened.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelPowerQueryMetadata", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelPowerQueryMetadata", "ExcelQueryMetadata")]
[OutputType(typeof(PSObject))]
public sealed class AddOfficeExcelPowerQueryMetadataCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetContext = "Context";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook document to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Connection name stored in workbook metadata.</summary>
    [Parameter(Mandatory = true)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Worksheet that should own query-table metadata. Defaults to the current DSL sheet when used inside New-OfficeExcel.</summary>
    [Parameter]
    [Alias("Sheet", "SheetName", "Worksheet")]
    public string? WorksheetName { get; set; }

    /// <summary>Optional query-table name.</summary>
    [Parameter]
    public string? QueryTableName { get; set; }

    /// <summary>Connection description.</summary>
    [Parameter]
    public string? Description { get; set; }

    /// <summary>Power Query M expression stored as metadata.</summary>
    [Parameter]
    public string? CommandText { get; set; }

    /// <summary>Request refresh-on-open metadata.</summary>
    [Parameter]
    public SwitchParameter RefreshOnOpen { get; set; }

    /// <summary>Emit metadata about the authored package parts.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {
            return;
        }

        string? sheetName = WorksheetName;
        if (string.IsNullOrWhiteSpace(sheetName) && string.Equals(ParameterSetName, ParameterSetContext, System.StringComparison.OrdinalIgnoreCase)) {
            sheetName = ExcelDslContext.Require(this).RequireSheet().Name;
        }

        var result = workbook.Document.AddPowerQueryMetadata(new ExcelPowerQueryMetadataOptions {
            Name = Name,
            WorksheetName = sheetName,
            QueryTableName = QueryTableName,
            Description = Description,
            CommandText = CommandText,
            RefreshOnOpen = RefreshOnOpen.IsPresent
        });
        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            var output = new PSObject();
            output.Properties.Add(new PSNoteProperty("Path", workbook.Document.FilePath));
            output.Properties.Add(new PSNoteProperty("ConnectionName", result.ConnectionName));
            output.Properties.Add(new PSNoteProperty("ConnectionId", result.ConnectionId));
            output.Properties.Add(new PSNoteProperty("QueryTableName", result.QueryTableName));
            output.Properties.Add(new PSNoteProperty("AddedWorkbookConnection", result.AddedWorkbookConnection));
            output.Properties.Add(new PSNoteProperty("AddedWorksheetQueryTable", result.AddedWorksheetQueryTable));
            output.Properties.Add(new PSNoteProperty("RefreshOnOpen", result.RefreshOnOpen));
            WriteObject(output);
        }
    }
}
#pragma warning restore CS1591