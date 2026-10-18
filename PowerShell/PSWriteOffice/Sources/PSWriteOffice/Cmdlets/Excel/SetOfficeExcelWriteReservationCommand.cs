using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets workbook write-reservation metadata.</summary>
/// <example>
///   <summary>Recommend read-only opening for a shared workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficeExcelWriteReservation -Path .\Report.xlsx -ReadOnlyRecommended -UserName 'Reporting Team' -PassThru |
///     Format-List ReadOnlyRecommended, UserName</code>
///   <para>Writes Excel file-sharing/write-reservation metadata without encrypting the file or protecting workbook structure.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelWriteReservation", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelWriteReservationSet")]
public sealed class SetOfficeExcelWriteReservationCommand : PSCmdlet
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

    /// <summary>Recommend opening the workbook as read-only.</summary>
    [Parameter]
    public SwitchParameter ReadOnlyRecommended { get; set; }

    /// <summary>User name stored in the write-reservation metadata.</summary>
    [Parameter]
    public string? UserName { get; set; }

    /// <summary>Optional write-reservation password. This is legacy Excel metadata, not package encryption.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>Optional precomputed legacy write-reservation hash.</summary>
    [Parameter]
    public string? LegacyPasswordHash { get; set; }

    /// <summary>Emit the updated write-reservation metadata.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        using var scope = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path ?? string.Empty, Document, readOnly: false);
        if (!ShouldProcess("Excel workbook", "Set write-reservation metadata"))
        {
            return;
        }

        scope.Document.SetWriteReservation(new ExcelWorkbookWriteReservationOptions {
            ReadOnlyRecommended = ReadOnlyRecommended.IsPresent,
            UserName = UserName,
            Password = Password,
            LegacyPasswordHash = LegacyPasswordHash
        });
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
