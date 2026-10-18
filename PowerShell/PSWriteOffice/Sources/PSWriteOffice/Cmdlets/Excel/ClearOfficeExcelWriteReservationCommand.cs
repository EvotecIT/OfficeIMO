using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Clears workbook write-reservation metadata.</summary>
/// <example>
///   <summary>Remove read-only recommendation metadata from a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Clear-OfficeExcelWriteReservation -Path .\Report.xlsx</code>
///   <para>Removes the workbook file-sharing/write-reservation node while leaving workbook protection and encryption state unchanged.</para>
/// </example>
[Cmdlet(VerbsCommon.Clear, "OfficeExcelWriteReservation", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelWriteReservationClear")]
public sealed class ClearOfficeExcelWriteReservationCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "InputPath", "FullName")]
    public string? Path { get; set; }

    /// <summary>Open workbook document to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Emit the resulting write-reservation metadata.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        using var scope = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path ?? string.Empty, Document, readOnly: false);
        if (!ShouldProcess("Excel workbook", "Clear write-reservation metadata"))
        {
            return;
        }

        scope.Document.ClearWriteReservation();
        scope.SaveIfOwned();

        if (PassThru.IsPresent)
        {
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
}
