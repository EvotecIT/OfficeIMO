using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Rtf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Creates an RTF document with plain paragraph content.</summary>
/// <example>
///   <summary>Create a small RTF file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$file = New-OfficeRtf -Path .\Report.rtf -Text 'Summary', 'Ready for review' -PassThru
/// Get-OfficeRtf -Path $file.FullName</code>
///   <para>Creates an RTF document with two paragraphs and returns the file.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeRtf", SupportsShouldProcess = true)]
[Alias("OfficeRtf", "RtfNew")]
[OutputType(typeof(FileInfo), typeof(RtfDocument))]
public sealed class NewOfficeRtfCommand : PSCmdlet {
    /// <summary>Destination path for the RTF file.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath", "OutputPath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Plain paragraph text to add to the document.</summary>
    [Parameter(Position = 1, ValueFromPipeline = true)]
    public string[]? Text { get; set; }

    /// <summary>Emit a FileInfo for chaining.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Return the OfficeIMO RTF document without saving.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

    private readonly List<string> _paragraphs = new();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Text != null) {
            _paragraphs.AddRange(Text);
        }
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        var document = RtfDocument.Create();
        foreach (var paragraph in _paragraphs) {
            document.AddParagraph(paragraph);
        }

        if (NoSave.IsPresent) {
            WriteObject(document);
            return;
        }

        var path = PdfCommandUtilities.ResolvePath(this, Path);
        if (!PdfCommandUtilities.ShouldWrite(this, path, "Write new RTF document")) {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(path);
        document.Save(path, encoding: new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        if (PassThru.IsPresent) {
            WriteObject(new FileInfo(path));
        }
    }
}
