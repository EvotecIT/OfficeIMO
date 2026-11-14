using System.Management.Automation;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Reports the media, paper sources, duplex, color, resolution and copy capabilities of a printer queue or IPP printer. Requires PowerShell 7.4 or newer.</summary>
/// <para>Name accepts an installed queue name or an ipp:// or ipps:// printer URI. Reading capabilities does not change printer settings.</para>
/// <example><summary>Read the capabilities of a queue and of an IPP printer.</summary><prefix>PS&gt; </prefix>
/// <code>Get-OfficePrinterCapability -Name 'Office printer'
/// Get-OfficePrinterCapability -Name 'ipp://printer.local/ipp/print'</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePrinterCapability", DefaultParameterSetName = "Default")]
#if !FRAMEWORK
[OutputType(typeof(PdfPrinterCapabilities))]
#endif
public sealed class GetOfficePrinterCapabilityCommand : AsyncPSCmdlet {
    /// <summary>Printer queue name or IPP printer URI.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipelineByPropertyName = true)]
    [ValidateNotNullOrEmpty]
    public string Name { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
#if FRAMEWORK
        await Task.CompletedTask;
        throw new PlatformNotSupportedException("Printer discovery requires PowerShell 7.4 or newer.");
#else
        WriteObject(await new PdfPrinterService().GetCapabilitiesAsync(Name, CancelToken));
#endif
    }
}
