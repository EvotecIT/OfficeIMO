using System;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Prepares PDF sheets and submits them to an explicitly named printer. Requires PowerShell 7.4 or newer.</summary>
/// <para>The returned receipt proves queue acceptance. It does not prove physical delivery. Check the queue after an interrupted submission before retrying.</para>
/// <example><summary>Print selected PDF pages as two pages per sheet.</summary><prefix>PS&gt; </prefix>
/// <code>Send-OfficePdfPrinter -Path .\Report.pdf -PrinterName 'Office printer' -Pages '1-3' -PagesPerSheet 2 -Duplex LongEdge</code></example>
[Cmdlet(VerbsCommunications.Send, "OfficePdfPrinter", SupportsShouldProcess = true)]
#if !FRAMEWORK
[OutputType(typeof(PdfPrintSubmission))]
#endif
public sealed class SendOfficePdfPrinterCommand : AsyncPSCmdlet {
    /// <summary>Local PDF source.</summary>
    [Parameter(Mandatory = true, Position = 0)] public string Path { get; set; } = string.Empty;
    /// <summary>Installed queue from Get-OfficePrinter.</summary>
    [Parameter(Mandatory = true)] public string PrinterName { get; set; } = string.Empty;
    /// <summary>Page selection such as 1-3,last.</summary>
    [Parameter] public string? Pages { get; set; }
    /// <summary>Source pages per sheet.</summary>
    [Parameter] [ValidateSet("1", "2", "4")] public int PagesPerSheet { get; set; } = 1;
    /// <summary>Copies submitted to the queue.</summary>
    [Parameter] [ValidateRange(1, 100)] public int Copies { get; set; } = 1;
    /// <summary>Duplex setting.</summary>
    [Parameter] [ValidateSet("PrinterDefault", "SingleSided", "LongEdge", "ShortEdge")]
    public string Duplex { get; set; } = "PrinterDefault";
    /// <summary>Paper source identifier reported for this queue.</summary>
    [Parameter] public string? PaperSourceId { get; set; }
    /// <summary>New destination file for a Windows file printer.</summary>
    [Parameter] public string? OutputFilePath { get; set; }
    /// <summary>Prepared sheet resolution.</summary>
    [Parameter] [ValidateRange(72d, 600d)] public double Dpi { get; set; } = 150;
    /// <summary>Printable margin in points.</summary>
    [Parameter] [ValidateRange(0d, double.MaxValue)] public double Margin { get; set; } = 18;
    /// <summary>Paper orientation.</summary>
    [Parameter] [ValidateSet("Automatic", "Portrait", "Landscape")]
    public string Orientation { get; set; } = "Automatic";
    /// <summary>Scaling of source pages.</summary>
    [Parameter] [ValidateSet("Fit", "ActualSize", "Fill")]
    public string ScaleMode { get; set; } = "Fit";
    /// <summary>PDF paper size, default A4.</summary>
    [Parameter] public PageSize PaperSize { get; set; } = PageSizes.A4;
    /// <summary>Password for an encrypted PDF; printing permissions remain enforced.</summary>
    [Parameter] public string? Password { get; set; }
    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
#if FRAMEWORK
        await Task.CompletedTask;
        throw new PlatformNotSupportedException("Printer delivery requires PowerShell 7.4 or newer.");
#else
        string input = PdfCommandUtilities.ResolveExistingFilePath(this, Path);
        string? output = OutputFilePath == null ? null : PdfCommandUtilities.ResolvePath(this, OutputFilePath);
        if (!ShouldProcess(PrinterName, "Submit prepared PDF sheets to the printer queue")) return;
        PdfDocument document = await PdfDocument.LoadAsync(input, new PdfLoadOptions { Password = Password }, cancellationToken: CancelToken);
        PdfPreparedPrintDocument sheets = PdfPrintRenderer.Prepare(document, new PdfPrintPlanRequest {
            InputPath = input, Pages = Pages, PagesPerSheet = PagesPerSheet, Margin = Margin,
            Orientation = (PdfPrintOrientation)Enum.Parse(typeof(PdfPrintOrientation), Orientation, true),
            ScaleMode = (PdfPrintScaleMode)Enum.Parse(typeof(PdfPrintScaleMode), ScaleMode, true), PaperSize = PaperSize
        }, new PdfPrintRenderOptions { Dpi = Dpi }, CancelToken);
        foreach (string diagnostic in sheets.Diagnostics) WriteWarning(diagnostic);
        PdfPrintSubmission receipt = await new PdfPrinterService().SubmitAsync(sheets, new PdfPrintDeliveryOptions {
            PrinterName = PrinterName, DocumentName = System.IO.Path.GetFileName(input), Copies = Copies,
            Duplex = (PdfPrintDuplex)Enum.Parse(typeof(PdfPrintDuplex), Duplex, true), PaperSourceId = PaperSourceId, OutputFilePath = output
        }, CancelToken);
        if (receipt.CleanupWarning != null) WriteWarning(receipt.CleanupWarning);
        WriteObject(receipt);
#endif
    }
}
