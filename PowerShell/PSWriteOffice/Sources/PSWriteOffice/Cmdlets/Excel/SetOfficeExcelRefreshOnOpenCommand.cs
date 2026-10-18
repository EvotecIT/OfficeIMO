using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures workbook data refresh metadata for Excel-compatible applications to run when the file opens.</summary>
/// <example>
///   <summary>Refresh pivot caches when the workbook opens.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$refresh = Set-OfficeExcelRefreshOnOpen -Path .\Report.xlsx -PivotTables -Connections -NoSavePivotSourceData -PassThru
/// Get-OfficeExcelDataModel -Path .\Report.xlsx |
///     Select-Object ConnectionPartCount, QueryTablePartCount</code>
///   <para>Sets workbook metadata through OfficeIMO so pivot caches refresh on open.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelRefreshOnOpen", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelRefreshOnOpen")]
[OutputType(typeof(PSObject))]
public sealed class SetOfficeExcelRefreshOnOpenCommand : PSCmdlet {
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

    /// <summary>Update pivot cache refresh-on-open metadata.</summary>
    [Parameter]
    public SwitchParameter PivotTables { get; set; }

    /// <summary>Update workbook connection refresh-on-open metadata.</summary>
    [Parameter]
    public SwitchParameter Connections { get; set; }

    /// <summary>Disable refresh-on-open instead of enabling it.</summary>
    [Parameter]
    public SwitchParameter Disable { get; set; }

    /// <summary>Preserve pivot cache source data.</summary>
    [Parameter]
    public SwitchParameter SavePivotSourceData { get; set; }

    /// <summary>Do not save pivot cache source data.</summary>
    [Parameter]
    public SwitchParameter NoSavePivotSourceData { get; set; }

    /// <summary>Emit the update summary.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {
            return;
        }


        bool targetPivotTables = PivotTables.IsPresent || !Connections.IsPresent;
        bool targetConnections = Connections.IsPresent || !PivotTables.IsPresent;
        bool? savePivotSourceData = ResolveSavePivotSourceData();
        ExcelRefreshOnOpenResult result = workbook.Document.SetRefreshOnOpen(
            enabled: !Disable.IsPresent,
            pivotTables: targetPivotTables,
            connections: targetConnections,
            savePivotSourceData: savePivotSourceData);

        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            var output = new PSObject();
            output.Properties.Add(new PSNoteProperty("Enabled", result.Enabled));
            output.Properties.Add(new PSNoteProperty("PivotCacheCount", result.PivotCacheCount));
            output.Properties.Add(new PSNoteProperty("ConnectionCount", result.ConnectionCount));
            WriteObject(output);
        }
    }

    private bool? ResolveSavePivotSourceData() {
        if (SavePivotSourceData.IsPresent && NoSavePivotSourceData.IsPresent) {
            throw new PSArgumentException("Specify either SavePivotSourceData or NoSavePivotSourceData, not both.");
        }

        if (SavePivotSourceData.IsPresent) {
            return true;
        }

        if (NoSavePivotSourceData.IsPresent) {
            return false;
        }

        return null;
    }
}