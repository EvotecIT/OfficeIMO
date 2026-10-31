using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Saves a live Markdown document using its native save policy and keeps it available.</summary>
/// <example><summary>Save to the document's associated path.</summary>
/// <code>$document | Save-OfficeMarkdown -PassThru</code></example>
/// <example><summary>Write a CRLF copy.</summary>
/// <code>Save-OfficeMarkdown -Document $document -Path ./Publish/Report.md -LineEnding CRLF -PassThru</code></example>
[Cmdlet(VerbsData.Save, "OfficeMarkdown", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class SaveOfficeMarkdownCommand : OfficeCmdlet, IMarkdownWriteOptionSource {
    /// <summary>Markdown document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public MarkdownDocument Document { get; set; } = null!;
    /// <summary>Optional destination; omission saves to the associated path.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }
    /// <summary>Native save settings, or a hashtable of native option properties.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(MarkdownSaveOptions)), ValidateNotNull]
    public MarkdownSaveOptions? WriteOptions { get; set; }
    /// <summary>Friendly Markdown writer profile.</summary>
    [Parameter]
    public OfficeMarkdownWriteProfile? WriteProfile { get; set; }
    /// <summary>Controls Markdown image serialization.</summary>
    [Parameter]
    public MarkdownImageRenderingMode? ImageRenderingMode { get; set; }
    /// <summary>Output line endings.</summary>
    [Parameter]
    public OfficeLineEnding? LineEnding { get; set; }
    /// <summary>Unordered list marker.</summary>
    [Parameter, ValidateSet("-", "*", "+")]
    public string? UnorderedListMarker { get; set; }
    /// <summary>Accept explicit native-save omissions, overriding WriteOptions when supplied.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }
    /// <summary>Return the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }
    string? IMarkdownWriteOptionSource.LineEnding => LineEnding?.ToString();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        string? destination = Path ?? Document.FilePath;
        if (string.IsNullOrWhiteSpace(destination)) {
            InvalidArgument(new PSArgumentException("Supply Path when the document has no associated destination."), Document);
            return;
        }
        MarkdownSaveOptions options = MarkdownOptionUtilities.BuildWriteOptions(this) ?? new MarkdownSaveOptions();
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowLoss))) options.AllowLoss = AllowLoss.IsPresent;
        destination = ResolveOutputPath(destination!);
        if (!ShouldProcess(destination, "Save Markdown document")) return;
        try {
            Document.Save(destination, options);
            if (PassThru) WriteObject(new FileInfo(destination));
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) { ItemError(error, "native-save-failed", destination); }
    }
}
