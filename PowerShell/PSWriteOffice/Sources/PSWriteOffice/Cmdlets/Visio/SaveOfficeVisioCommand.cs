using System.IO;
using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Saves an OfficeIMO.Visio document.</summary>
/// <example>
///   <summary>Save a loaded diagram under a new path.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$diagram = Get-OfficeVisio -Path .\ServiceMap.vsdx
/// $diagram | Save-OfficeVisio -Path .\ServiceMap-copy.vsdx -PassThru</code>
///   <para>Saves an existing OfficeIMO.Visio document to another .vsdx file.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeVisio", SupportsShouldProcess = true)]
[Alias("VisioSave")]
[OutputType(typeof(VisioDocument))]
public sealed class SaveOfficeVisioCommand : PSCmdlet {
    /// <summary>Visio document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public VisioDocument Document { get; set; } = null!;

    /// <summary>Optional save-as path.</summary>
    [Parameter(Position = 1)]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit the document object for further processing.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (string.IsNullOrWhiteSpace(Path)) {
            var associatedPath = Document.FilePath;
            if (string.IsNullOrWhiteSpace(associatedPath)) {
                throw new PSInvalidOperationException("No file path provided. Use -Path or load the document from disk.");
            }

            var targetPath = associatedPath!;

            if (!ShouldProcess(targetPath, "Save Visio document")) {
                return;
            }

            Document.Save();
            if (Open.IsPresent) {
                FileOpenService.Open(targetPath);
            }

            if (PassThru.IsPresent) {
                WriteObject(Document);
            }
            return;
        }

        var fullPath = VisioCommandUtilities.ResolvePath(this, Path!);
        if (!ShouldProcess(fullPath, "Save Visio document")) {
            return;
        }

        VisioCommandUtilities.EnsureDirectory(fullPath);
        Document.Save(fullPath);

        if (Open.IsPresent) {
            FileOpenService.Open(fullPath);
        }

        if (PassThru.IsPresent) {
            WriteObject(Document);
        }
    }
}