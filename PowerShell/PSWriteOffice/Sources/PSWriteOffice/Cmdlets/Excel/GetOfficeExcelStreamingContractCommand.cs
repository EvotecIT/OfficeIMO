#pragma warning disable CS1591
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Reports large-workbook streaming and direct-writer suitability.</summary>
/// <example>
///   <summary>Choose the write path for a large export.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$contract = Get-OfficeExcelStreamingContract -Path .\LargeExport.xlsx
/// [pscustomobject]@{
///     EstimatedCells = $contract.EstimatedCellCount
///     DirectWriter   = $contract.HasDirectDataSetFastSaveState
///     Recommendation = $contract.Recommendation
/// }</code>
///   <para>Reports whether the workbook is already using OfficeIMO direct tabular state and gives a size-based recommendation for future imports or exports.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelStreamingContract", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelStreamingContract")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelStreamingContractCommand : PSCmdlet {
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
        var report = workbook.Document.GetStreamingContract();
        var output = new PSObject();
        output.Properties.Add(new PSNoteProperty("Path", workbook.Document.FilePath));
        output.Properties.Add(new PSNoteProperty("WorksheetCount", report.WorksheetCount));
        output.Properties.Add(new PSNoteProperty("EstimatedCellCount", report.EstimatedCellCount));
        output.Properties.Add(new PSNoteProperty("HasDirectDataSetFastSaveState", report.HasDirectDataSetFastSaveState));
        output.Properties.Add(new PSNoteProperty("HasDeferredDirectDataSetImport", report.HasDeferredDirectDataSetImport));
        output.Properties.Add(new PSNoteProperty("Recommendation", report.Recommendation));
        WriteObject(output);
    }
}
#pragma warning restore CS1591