using OfficeIMO.Core.Internal;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Saves a Markdown document without changing its lifetime.</summary>
/// <example>
///   <summary>Save a Markdown document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc | Save-OfficeMarkdown -Path .\Report.md</code>
///   <para>Writes the Markdown artifact and keeps the document available for further changes.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeMarkdown", SupportsShouldProcess = true)]
[OutputType(typeof(MarkdownDocument))]
public sealed class SaveOfficeMarkdownCommand : PSCmdlet
    , IMarkdownWriteOptionSource {
    /// <summary>Markdown document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public MarkdownDocument Document { get; set; } = null!;

    /// <summary>Destination Markdown path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional Markdown writer options.</summary>
    [Parameter]
    public MarkdownSaveOptions? WriteOptions { get; set; }

    /// <summary>Friendly Markdown writer profile.</summary>
    [Parameter]
    public OfficeMarkdownWriteProfile? WriteProfile { get; set; }

    /// <summary>Controls how Markdown images are serialized.</summary>
    [Parameter]
    public MarkdownImageRenderingMode? ImageRenderingMode { get; set; }

    /// <summary>Markdown line ending: CRLF, LF, CR, or a literal line ending string.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficeLineEndingArgumentCompleter))]
    public string? LineEnding { get; set; }

    /// <summary>Unordered list marker: '-', '*', or '+'.</summary>
    [Parameter]
    public string? UnorderedListMarker { get; set; }

    /// <summary>Emit the Markdown document rather than the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var fullPath = PdfCommandUtilities.ResolvePath(this, Path);
        if (!PdfCommandUtilities.ShouldWrite(this, fullPath, "Save Markdown document")) {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(fullPath);
        OfficeFileOutput.Write(fullPath, overwrite: true, temporaryPath => File.WriteAllText(temporaryPath, Document.ToMarkdown(MarkdownOptionUtilities.BuildWriteOptions(this)), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)));

        if (PassThru.IsPresent) {
            WriteObject(Document);
        }
    }

}
