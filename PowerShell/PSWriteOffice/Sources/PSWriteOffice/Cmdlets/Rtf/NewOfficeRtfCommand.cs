using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;
using OfficeIMO.Rtf;
namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Creates an RTF document, optionally with paragraphs from -Text or the pipeline.</summary>
/// <example><summary>Use the native OfficeRtf command.</summary><code>New-OfficeRtf -Path .\Note.rtf -Text 'Server maintenance starts at 22:00.'</code></example>
[Cmdlet("New", "OfficeRtf", DefaultParameterSetName = "Live", SupportsShouldProcess = true)]
[OutputType(typeof(RtfDocument))]
public sealed class NewOfficeRtfCommand : OfficeCmdlet {
    /// <summary>Path for this operation.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Content", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Live", Position = 0)]
    public string? Path { get; set; }
    /// <summary>Text for this operation.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true, Position = 1, ValueFromPipeline = true)]
    public string[]? Text { get; set; }
    /// <summary>Force for this operation.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "Live")]
    public SwitchParameter Force { get; set; }
    /// <summary>PassThru for this operation.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public SwitchParameter PassThru { get; set; }
    /// <summary>Show for this operation.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public SwitchParameter Show { get; set; }
    /// <summary>Content for this operation.</summary>
    [Parameter(ParameterSetName = "Content", Mandatory = true, Position = 1)]
    public ScriptBlock? Content { get; set; }

    private readonly List<string> _paragraphs = new();
    /// <inheritdoc />
    protected override void ProcessRecord() { if (Text != null) _paragraphs.AddRange(Text); }
    /// <inheritdoc />
    protected override void EndProcessing() {
        string? output = Path == null ? null : OfficePath.ResolveOutput(SessionState, Path, false);
        if (output != null) OfficePath.ValidateDestination(output, !Force);
        if ((Content != null || Text != null) && !ShouldProcess(output!, "Create RTF document")) return;
        var document = output == null ? RtfDocument.Create() : RtfDocument.Create(output);
        if (Content == null && Text == null) { WriteObject(document); return; }
        try {
            if (Content != null) { using (OfficeDocumentContext.Push(document)) Content.InvokeReturnAsIs(); }
            else foreach (string paragraph in _paragraphs) document.AddParagraph(paragraph);
            PdfCommandUtilities.EnsureDirectory(output!);
            OfficeIMO.Core.Internal.OfficeFileOutput.Write(output!, Force, temporary => document.Save(temporary));
        } finally { document.Dispose(); }
        if (Show) OfficeShow.Open(output!);
        if (PassThru) WriteObject(new FileInfo(output!));
    }
}
