using System.Management.Automation;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Reports the state of a print job submitted with Send-OfficePrinter, optionally waiting until it finishes. Requires PowerShell 7.4 or newer.</summary>
/// <para>States are the same for every transport: Pending, Held, Processing, Stopped, Canceled, Aborted, Completed and Removed. Removed means the queue no longer reports the job: it finished and was purged, was deleted, or never existed.</para>
/// <example><summary>Wait up to two minutes for a submitted job.</summary><prefix>PS&gt; </prefix>
/// <code>$receipt = Send-OfficePrinter -Path .\Report.pdf -PrinterName 'Office printer'
/// $receipt | Get-OfficePrintJob -Wait -TimeoutSeconds 120</code></example>
/// <example><summary>Check a job by printer and identifier.</summary><prefix>PS&gt; </prefix>
/// <code>Get-OfficePrintJob -PrinterName 'Office printer' -JobId 'Office_printer-42'</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePrintJob", DefaultParameterSetName = "Default")]
#if !FRAMEWORK
[OutputType(typeof(PdfPrintJobStatus))]
#endif
public sealed class GetOfficePrintJobCommand : AsyncPSCmdlet {
    /// <summary>The printer queue name or IPP URI from the receipt.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipelineByPropertyName = true)]
    [ValidateNotNullOrEmpty]
    public string PrinterName { get; set; } = string.Empty;

    /// <summary>The job identifier from the receipt.</summary>
    [Parameter(Mandatory = true, Position = 1, ValueFromPipelineByPropertyName = true)]
    [ValidateNotNullOrEmpty]
    public string JobId { get; set; } = string.Empty;

    /// <summary>Wait until the job is canceled, aborted, completed or removed.</summary>
    [Parameter] public SwitchParameter Wait { get; set; }

    /// <summary>Longest wait in seconds when Wait is set.</summary>
    [Parameter] [ValidateRange(1, 86400)] public int TimeoutSeconds { get; set; } = 300;

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
#if FRAMEWORK
        await Task.CompletedTask;
        throw new PlatformNotSupportedException("Print job tracking requires PowerShell 7.4 or newer.");
#else
        var service = new PdfPrinterService();
        WriteObject(Wait
            ? await service.WaitForJobAsync(PrinterName, JobId, TimeSpan.FromSeconds(TimeoutSeconds), cancellationToken: CancelToken)
            : await service.GetJobAsync(PrinterName, JobId, CancelToken));
#endif
    }
}
