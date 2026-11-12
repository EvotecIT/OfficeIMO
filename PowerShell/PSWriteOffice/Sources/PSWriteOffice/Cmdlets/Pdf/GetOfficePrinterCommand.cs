using System.Management.Automation;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Lists system printer queues, or the capabilities or paper sources of a named queue. Requires PowerShell 7.4 or newer.</summary>
/// <example><summary>Discover queues, capabilities and trays.</summary><prefix>PS&gt; </prefix>
/// <code>Get-OfficePrinter
/// Get-OfficePrinter -Name 'Office printer' -Capability
/// Get-OfficePrinter -Name 'Office printer' -PaperSource</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePrinter", DefaultParameterSetName = "Printer")]
#if !FRAMEWORK
[OutputType(typeof(PdfPrinterInfo), typeof(PdfPaperSourceInfo), typeof(PdfPrinterCapabilities))]
#endif
public sealed class GetOfficePrinterCommand : AsyncPSCmdlet {
    /// <summary>Printer name patterns; Capability and PaperSource require one exact printer name.</summary>
    [Parameter(Position = 0, ParameterSetName = "Printer")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PaperSource")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Capability")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public string[] Name { get; set; } = Array.Empty<string>();

    /// <summary>List paper sources of the single printer selected by Name.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "PaperSource")]
    public SwitchParameter PaperSource { get; set; }

    /// <summary>Return the media, trays, duplex, color, resolution, collation and copy capabilities of the single printer selected by Name.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Capability")]
    public SwitchParameter Capability { get; set; }
    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
#if FRAMEWORK
        await Task.CompletedTask;
        throw new PlatformNotSupportedException("Printer discovery requires PowerShell 7.4 or newer.");
#else
        var service = new PdfPrinterService();
        if (PaperSource || Capability) {
            if (Name.Length != 1 || WildcardPattern.ContainsWildcardCharacters(Name[0]))
                throw new PSArgumentException((PaperSource ? "PaperSource" : "Capability") + " requires one exact printer name.", nameof(Name));
            PdfPrinterCapabilities capabilities = await service.GetCapabilitiesAsync(Name[0], CancelToken);
            if (Capability) WriteObject(capabilities);
            else foreach (var source in capabilities.PaperSources) WriteObject(source);
        } else {
            var patterns = Name.Select(name => new WildcardPattern(name, WildcardOptions.IgnoreCase)).ToArray();
            foreach (var printer in await service.GetPrintersAsync(CancelToken))
                if (patterns.Length == 0 || patterns.Any(pattern => pattern.IsMatch(printer.Name))) WriteObject(printer);
        }
#endif
    }
}
