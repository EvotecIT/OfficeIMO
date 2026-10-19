using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets workbook write-reservation metadata.</summary>
/// <example>
///   <summary>Inspect whether a workbook recommends read-only opening.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeExcelWriteReservation -Path .\Report.xlsx |
///     Format-List Exists, ReadOnlyRecommended, UserName, HasPasswordHash</code>
///   <para>Reports Excel file-sharing/write-reservation metadata separately from workbook protection and package encryption.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelWriteReservation", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelWriteReservation")]
public sealed class GetOfficeExcelWriteReservationCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path to inspect.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "InputPath", "FullName")]
    public string? Path { get; set; }

    /// <summary>Open workbook document to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        using var scope = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path ?? string.Empty, Document, readOnly: true);
        var info = scope.Document.GetWriteReservation();
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Exists", info.Exists));
        record.Properties.Add(new PSNoteProperty("ReadOnlyRecommended", info.ReadOnlyRecommended));
        record.Properties.Add(new PSNoteProperty("UserName", info.UserName));
        record.Properties.Add(new PSNoteProperty("LegacyPasswordHash", info.LegacyPasswordHash));
        record.Properties.Add(new PSNoteProperty("HasPasswordHash", info.HasPasswordHash));
        WriteObject(record);
    }
}
