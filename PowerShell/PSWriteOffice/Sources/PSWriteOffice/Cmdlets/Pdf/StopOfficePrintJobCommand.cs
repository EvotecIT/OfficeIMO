using System.Management.Automation;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Cancels a print job submitted with Send-OfficePrinter and returns its state afterwards. Requires PowerShell 7.4 or newer.</summary>
/// <para>A job that has already finished is reported unchanged. Pages the printer has already received may still print.</para>
/// <example><summary>Cancel a submitted job.</summary><prefix>PS&gt; </prefix>
/// <code>Send-OfficePrinter -Path .\Report.pdf -PrinterName 'Office printer' | Stop-OfficePrintJob</code></example>
[Cmdlet(VerbsLifecycle.Stop, "OfficePrintJob", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Medium, DefaultParameterSetName = "Default")]
#if !FRAMEWORK
[OutputType(typeof(PdfPrintJobStatus))]
#endif
public sealed class StopOfficePrintJobCommand : AsyncPSCmdlet {
    /// <summary>The printer queue name or IPP URI from the receipt.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipelineByPropertyName = true)]
    [ValidateNotNullOrEmpty]
    public string PrinterName { get; set; } = string.Empty;

    /// <summary>The job identifier from the receipt.</summary>
    [Parameter(Mandatory = true, Position = 1, ValueFromPipelineByPropertyName = true)]
    [ValidateNotNullOrEmpty]
    public string JobId { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
#if FRAMEWORK
        await Task.CompletedTask;
        throw new PlatformNotSupportedException("Print job tracking requires PowerShell 7.4 or newer.");
#else
        if (!ShouldProcess(PrinterName, "Cancel print job " + JobId)) return;
        WriteObject(await new PdfPrinterService().CancelJobAsync(PrinterName, JobId, CancelToken));
#endif
    }
}
