using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;
using OfficeIMO.Rtf;
namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Saves an RtfDocument to its original path or to a new path.</summary>
/// <example><summary>Use the native OfficeRtf command.</summary><code>Save-OfficeRtf -Document $rtf.Document -Path .\Copy.rtf</code></example>
[Cmdlet("Save", "OfficeRtf", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class SaveOfficeRtfCommand : OfficeCmdlet {
    /// <summary>Document for this operation.</summary>
    [Parameter(ParameterSetName = "Document", Mandatory = true, ValueFromPipeline = true)]
    public RtfDocument Document { get; set; } = null!;
    /// <summary>Path for this operation.</summary>
    [Parameter(ParameterSetName = "Document", Position = 0)]
    public string? Path { get; set; }
    /// <summary>PassThru for this operation.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        string destination = Path ?? Document!.FilePath ?? throw new PSArgumentException("Supply Path for a document without an associated destination.");
        string output = OfficePath.ResolveOutput(SessionState, destination, false);
        if (!ShouldProcess(output, "Save RTF document")) return;
        PdfCommandUtilities.EnsureDirectory(output);
        if (Path == null) Document!.Save(); else Document!.Save(output);
        if (PassThru) WriteObject(new FileInfo(output));
    }
}
