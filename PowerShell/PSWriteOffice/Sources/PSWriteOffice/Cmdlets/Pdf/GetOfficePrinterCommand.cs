using System;
using System.Management.Automation;
using System.Threading.Tasks;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Lists system printer queues, or the paper sources of a named queue. Requires PowerShell 7.4 or newer.</summary>
/// <example><summary>Discover queues and trays.</summary><prefix>PS&gt; </prefix>
/// <code>Get-OfficePrinter
/// Get-OfficePrinter -PaperSources 'Office printer'</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePrinter")]
#if !FRAMEWORK
[OutputType(typeof(PdfPrinterInfo), typeof(PdfPaperSourceInfo))]
#endif
public sealed class GetOfficePrinterCommand : AsyncPSCmdlet {
    /// <summary>Named printer whose paper sources should be listed.</summary>
    [Parameter] public string? PaperSources { get; set; }
    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
#if FRAMEWORK
        await Task.CompletedTask;
        throw new PlatformNotSupportedException("Printer discovery requires PowerShell 7.4 or newer.");
#else
        var service = new PdfPrinterService();
        if (PaperSources == null) foreach (var printer in await service.GetPrintersAsync(CancelToken)) WriteObject(printer);
        else foreach (var source in await service.GetPaperSourcesAsync(PaperSources, CancelToken)) WriteObject(source);
#endif
    }
}
