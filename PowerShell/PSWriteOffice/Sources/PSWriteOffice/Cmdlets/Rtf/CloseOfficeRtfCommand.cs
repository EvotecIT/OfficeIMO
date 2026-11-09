using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;
using OfficeIMO.Rtf;
namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Releases an RtfDocument without saving.</summary>
/// <example><summary>Use the native OfficeRtf command.</summary><code>Close-OfficeRtf -Document $rtf.Document</code></example>
[Cmdlet("Close", "OfficeRtf", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(void))]
public sealed class CloseOfficeRtfCommand : OfficeCmdlet {
    /// <summary>Document for this operation.</summary>
    [Parameter(ParameterSetName = "Document", Mandatory = true, ValueFromPipeline = true)]
    public RtfDocument Document { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!ShouldProcess(Document!.FilePath ?? "Unsaved RTF document", "Close RTF document without saving")) return;
        Document.Dispose();
    }
}
