using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif
using PSWriteOffice.Services.Pdf;
using System.Security;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Prints a document of any format OfficeIMO converts to PDF on an explicitly named printer queue or IPP printer. Requires PowerShell 7.4 or newer.</summary>
/// <para>PDF files are printed directly; other documents are converted to PDF in memory first, and conversion warnings are written as warnings. Send-OfficePdfPrinter remains an alias of this command. PrinterName accepts an installed queue name or an ipp:// or ipps:// printer URI.</para>
/// <para>Without PaperSize, sheets use the printer's default media. The returned receipt proves queue acceptance, not physical delivery. Follow it with Get-OfficePrintJob, and check the queue after an interrupted submission before retrying.</para>
/// <example><summary>Print selected PDF pages as two pages per sheet.</summary><prefix>PS&gt; </prefix>
/// <code>Send-OfficePrinter -Path .\Report.pdf -PrinterName 'Office printer' -Page '1-3' -PagesPerSheet 2 -Duplex LongEdge</code></example>
/// <example><summary>Print a Word document to an IPP printer and wait for it to finish.</summary><prefix>PS&gt; </prefix>
/// <code>Send-OfficePrinter -Path .\Letter.docx -PrinterName 'ipp://printer.local/ipp/print' | Get-OfficePrintJob -Wait</code></example>
/// <example><summary>Print odd pages four to a sheet in grayscale.</summary><prefix>PS&gt; </prefix>
/// <code>Send-OfficePrinter -Path .\Report.pdf -PrinterName 'Office printer' -PageSubset Odd -PagesPerSheet 4 -Grayscale</code></example>
[Cmdlet(VerbsCommunications.Send, "OfficePrinter", SupportsShouldProcess = true, DefaultParameterSetName = "Path")]
[Alias("Send-OfficePdfPrinter")]
#if !FRAMEWORK
[OutputType(typeof(PdfPrintSubmission))]
#endif
public sealed class SendOfficePrinterCommand : AsyncPSCmdlet {
    /// <summary>Local document to print: PDF or any format OfficeIMO converts to PDF.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path"), SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Literal document paths, including FileInfo.FullName.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath"), Alias("FullName", "PSPath")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Installed queue from Get-OfficePrinter, or an ipp:// or ipps:// printer URI.</summary>
    [Parameter(Mandatory = true)] public string PrinterName { get; set; } = string.Empty;
    /// <summary>Page selection such as 1-3,last.</summary>
    [Parameter, Alias("Pages"), OfficePageSelectorTransformation] public OfficePageSelector? Page { get; set; }
    /// <summary>Source pages per sheet, in row-major order.</summary>
    [Parameter] [ValidateSet("1", "2", "4", "6", "9")] public int PagesPerSheet { get; set; } = 1;
    /// <summary>Keep all selected pages, or only odd or even original page numbers.</summary>
    [Parameter] [ValidateSet("All", "Odd", "Even")] public string PageSubset { get; set; } = "All";
    /// <summary>Print in grayscale, regardless of the device's color support. IPP printers and CUPS queues reached over IPP that list monochrome printing receive the vector PDF and convert it to gray themselves; others receive gray raster sheets.</summary>
    [Parameter] public SwitchParameter Grayscale { get; set; }
    /// <summary>With Grayscale, always send gray raster sheets, so the printout matches the preview pixel for pixel. Slower than letting the printer convert vector PDF. Windows queues always receive raster sheets.</summary>
    [Parameter] public SwitchParameter RasterizeGrayscale { get; set; }
    /// <summary>Copies submitted to the queue.</summary>
    [Parameter] [ValidateRange(1, 100)] public int Copies { get; set; } = 1;
    /// <summary>Duplex setting.</summary>
    [Parameter] [ValidateSet("PrinterDefault", "SingleSided", "LongEdge", "ShortEdge")]
    public string Duplex { get; set; } = "PrinterDefault";
    /// <summary>Paper source identifier reported by Get-OfficePrinterCapability.</summary>
    [Parameter] public string? PaperSourceId { get; set; }
    /// <summary>New destination file for a Windows file printer.</summary>
    [Parameter, Alias("OutputFilePath")] public string? OutputPath { get; set; }
    /// <summary>Highest resolution used when sheets are rasterized for delivery.</summary>
    [Parameter] [ValidateRange(72d, 600d)] public double Dpi { get; set; } = 300;
    /// <summary>Printable margin in points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength Margin { get; set; } = OfficeLength.Points(18);
    /// <summary>Paper orientation. When omitted, each sheet takes the orientation that places its pages largest, so two-up and six-up sheets of portrait pages print in landscape.</summary>
    [Parameter]
    public OfficePageOrientation? Orientation { get; set; }
    /// <summary>Scaling of source pages.</summary>
    [Parameter] [ValidateSet("Fit", "ActualSize", "Fill")]
    public string ScaleMode { get; set; } = "Fit";
    /// <summary>Paper size. Defaults to the printer's default media, or A4 when the printer reports none.</summary>
    [Parameter, OfficePageSizeTransformation(typeof(OfficePageSize)), ArgumentCompleter(typeof(OfficePdfPageSizeArgumentCompleter))]
    public OfficePageSize PaperSize { get; set; } = OfficePageSizes.A4;
    /// <summary>Password for an encrypted PDF; printing permissions remain enforced.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
#if FRAMEWORK
        await Task.CompletedTask;
        throw new PlatformNotSupportedException("Printer delivery requires PowerShell 7.4 or newer.");
#else
        bool literal = ParameterSetName == "LiteralPath";
        var inputs = OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal);
        string? output = OutputPath == null ? null : PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (output != null && inputs.Count != 1) throw new PSArgumentException("A file printer OutputPath requires one document.");
        // Every input is checked before the first job is submitted, so an invalid request never leaves a partial print run.
        if (Password != null && inputs.Any(input => !string.Equals(System.IO.Path.GetExtension(input), ".pdf", StringComparison.OrdinalIgnoreCase)))
            throw new PSArgumentException("Password applies only to PDF documents.", nameof(Password));
        var work = new System.Collections.Generic.List<string>();
        foreach (string input in inputs)
            if (ShouldProcess(PrinterName, "Submit " + input + " as a print job")) work.Add(input);
        string? password = Password.ToPlainText();
        var service = new PdfPrinterService();
        foreach (string input in work) {
            OfficePrintResult result = await service.PrintAsync(new OfficePrintRequest {
                Layout = new PdfPrintPlanRequest {
                    InputPath = input, Pages = Page?.ToString(), PagesPerSheet = PagesPerSheet, Margin = Margin.ToPoints(),
                    Orientation = Orientation.HasValue ? Orientation.Value == OfficePageOrientation.Landscape ? PdfPrintOrientation.Landscape : PdfPrintOrientation.Portrait : PdfPrintOrientation.Automatic,
                    ScaleMode = (PdfPrintScaleMode)Enum.Parse(typeof(PdfPrintScaleMode), ScaleMode, true),
                    PageSubset = (PdfPrintPageSubset)Enum.Parse(typeof(PdfPrintPageSubset), PageSubset, true),
                    ColorMode = Grayscale ? PdfPrintColorMode.Grayscale : PdfPrintColorMode.Color,
                    PaperSize = new PdfPageSize(PaperSize.Width, PaperSize.Height),
                    PdfPassword = password
                },
                UsePrinterPaper = !MyInvocation.BoundParameters.ContainsKey(nameof(PaperSize)),
                Delivery = new PdfPrintDeliveryOptions {
                    PrinterName = PrinterName, DocumentName = System.IO.Path.GetFileName(input), Copies = Copies,
                    Duplex = (PdfPrintDuplex)Enum.Parse(typeof(PdfPrintDuplex), Duplex, true), PaperSourceId = PaperSourceId, OutputFilePath = output,
                    MaximumRasterDpi = Dpi, RasterizeGrayscale = RasterizeGrayscale
                }
            }, CancelToken);
            if (result.Conversion is { } conversion)
                foreach (OfficeDiagnostic diagnostic in conversion.Diagnostics)
                    if (diagnostic.Severity != OfficeDiagnosticSeverity.Information) WriteWarning(diagnostic.Message);
            foreach (string diagnostic in result.Diagnostics) WriteWarning(diagnostic);
            PdfPrintSubmission receipt = result.Submission;
            if (receipt.IgnoredAttributes.Count > 0) WriteWarning("The printer ignored: " + string.Join(", ", receipt.IgnoredAttributes) + ".");
            if (receipt.CleanupWarning != null) WriteWarning(receipt.CleanupWarning);
            WriteObject(receipt);
        }
#endif
    }
}
