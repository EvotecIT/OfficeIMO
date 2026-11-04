using System;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
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

/// <summary>Prepares PDF sheets and submits them to an explicitly named printer. Requires PowerShell 7.4 or newer.</summary>
/// <para>The returned receipt proves queue acceptance. It does not prove physical delivery. Check the queue after an interrupted submission before retrying.</para>
/// <example><summary>Print selected PDF pages as two pages per sheet.</summary><prefix>PS&gt; </prefix>
/// <code>Send-OfficePdfPrinter -Path .\Report.pdf -PrinterName 'Office printer' -Page '1-3' -PagesPerSheet 2 -Duplex LongEdge</code></example>
[Cmdlet(VerbsCommunications.Send, "OfficePdfPrinter", SupportsShouldProcess = true, DefaultParameterSetName = "Path")]
#if !FRAMEWORK
[OutputType(typeof(PdfPrintSubmission))]
#endif
public sealed class SendOfficePdfPrinterCommand : AsyncPSCmdlet {
    /// <summary>Local PDF source.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path"), SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Literal PDF sources, including FileInfo.FullName.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath"), Alias("FullName", "PSPath")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Installed queue from Get-OfficePrinter.</summary>
    [Parameter(Mandatory = true)] public string PrinterName { get; set; } = string.Empty;
    /// <summary>Page selection such as 1-3,last.</summary>
    [Parameter, PdfPageSelectorTransformation] public PdfPageSelector? Page { get; set; }
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
    [Parameter] public string? OutputPath { get; set; }
    /// <summary>Prepared sheet resolution.</summary>
    [Parameter] [ValidateRange(72d, 600d)] public double Dpi { get; set; } = 150;
    /// <summary>Printable margin in points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength Margin { get; set; } = OfficeLength.Points(18);
    /// <summary>Paper orientation.</summary>
    [Parameter]
    public OfficePageOrientation? Orientation { get; set; }
    /// <summary>Scaling of source pages.</summary>
    [Parameter] [ValidateSet("Fit", "ActualSize", "Fill")]
    public string ScaleMode { get; set; } = "Fit";
    /// <summary>PDF paper size, default A4.</summary>
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
        if (output != null && inputs.Count != 1) throw new PSArgumentException("A file printer OutputPath requires one PDF.");
        var work = new System.Collections.Generic.List<string>();
        foreach (string input in inputs)
            if (ShouldProcess(PrinterName, "Submit " + input + " as prepared PDF sheets")) work.Add(input);
        string? password = Password.ToPlainText();
        foreach (string input in work) {
            using PdfDocument document = await PdfDocument.LoadAsync(input, new PdfLoadOptions { Password = password }, cancellationToken: CancelToken);
            PdfPreparedPrintDocument sheets = PdfPrintRenderer.Prepare(document, new PdfPrintPlanRequest {
                InputPath = input, Pages = Page?.ToString(), PagesPerSheet = PagesPerSheet, Margin = Margin.ToPoints(),
                Orientation = Orientation.HasValue ? Orientation.Value == OfficePageOrientation.Landscape ? PdfPrintOrientation.Landscape : PdfPrintOrientation.Portrait : PdfPrintOrientation.Automatic,
                ScaleMode = (PdfPrintScaleMode)Enum.Parse(typeof(PdfPrintScaleMode), ScaleMode, true),
                PaperSize = new PdfPageSize(PaperSize.Width, PaperSize.Height)
            }, new PdfPrintRenderOptions { Dpi = Dpi }, CancelToken);
            foreach (string diagnostic in sheets.Diagnostics) WriteWarning(diagnostic);
            PdfPrintSubmission receipt = await new PdfPrinterService().SubmitAsync(sheets, new PdfPrintDeliveryOptions {
                PrinterName = PrinterName, DocumentName = System.IO.Path.GetFileName(input), Copies = Copies,
                Duplex = (PdfPrintDuplex)Enum.Parse(typeof(PdfPrintDuplex), Duplex, true), PaperSourceId = PaperSourceId, OutputFilePath = output
            }, CancelToken);
            if (receipt.CleanupWarning != null) WriteWarning(receipt.CleanupWarning);
            WriteObject(receipt);
        }
#endif
    }
}
