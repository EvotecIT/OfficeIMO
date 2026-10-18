#pragma warning disable CS1591
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Inspects workbook data model, query, connection, and external-link package parts.</summary>
/// <example>
///   <summary>Detect query/model metadata that OfficeIMO preserves but does not execute.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$model = Get-OfficeExcelDataModel -Path .\WorkbookWithQueries.xlsx
/// if ($model.HasDataModelOrQueries) {
///     $model.Details | ForEach-Object { "Preserved package part: $_" }
/// }</code>
///   <para>Separates preserved connection/query/data-model parts from executable refresh behavior so automation can decide whether Excel refresh is required.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelDataModel", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelDataModel", "ExcelPowerQuery")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelDataModelCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var report = workbook.Document.InspectDataModel();
        var output = new PSObject();
        output.Properties.Add(new PSNoteProperty("Path", workbook.Document.FilePath));
        output.Properties.Add(new PSNoteProperty("HasDataModelOrQueries", report.HasDataModelOrQueries));
        output.Properties.Add(new PSNoteProperty("ConnectionPartCount", report.ConnectionPartCount));
        output.Properties.Add(new PSNoteProperty("QueryTablePartCount", report.QueryTablePartCount));
        output.Properties.Add(new PSNoteProperty("ModelPartCount", report.ModelPartCount));
        output.Properties.Add(new PSNoteProperty("ExternalLinkPartCount", report.ExternalLinkPartCount));
        output.Properties.Add(new PSNoteProperty("Details", report.Details));
        WriteObject(output);
    }
}
#pragma warning restore CS1591