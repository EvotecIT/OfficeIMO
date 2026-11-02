using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Creates a live Markdown document, or atomically publishes a completed content block.</summary>
/// <example><summary>Create a report in one block.</summary>
/// <code>New-OfficeMarkdown -Path ./Report.md -PassThru { MarkdownHeading -Text 'Report' }</code></example>
/// <example><summary>Compose a live document before saving.</summary>
/// <code>$document = New-OfficeMarkdown -Path ./Report.md
/// $document | MarkdownHeading -Text 'Report'
/// $document | Save-OfficeMarkdown
/// $document | Close-OfficeMarkdown</code></example>
[Cmdlet(VerbsCommon.New, "OfficeMarkdown", DefaultParameterSetName = "Live", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(MarkdownDocument))]
public sealed class NewOfficeMarkdownCommand : OfficeNewLifecycleCmdlet<MarkdownDocument>, IMarkdownWriteOptionSource {
    /// <summary>Allow an existing destination; a failed block preserves its original contents.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }
    /// <summary>Native Markdown save settings or a hashtable, used when publishing a block.</summary>
    [Parameter(ParameterSetName = "Content"), OfficeOptionsTransformation(typeof(MarkdownSaveOptions)), ValidateNotNull]
    public MarkdownSaveOptions? WriteOptions { get; set; }
    /// <summary>Writer profile used when publishing a block.</summary>
    [Parameter(ParameterSetName = "Content")]
    public OfficeMarkdownWriteProfile? WriteProfile { get; set; }
    /// <summary>Image serialization used when publishing a block.</summary>
    [Parameter(ParameterSetName = "Content")]
    public MarkdownImageRenderingMode? ImageRenderingMode { get; set; }
    /// <summary>Line endings used when publishing a block.</summary>
    [Parameter(ParameterSetName = "Content")]
    public OfficeLineEnding? LineEnding { get; set; }
    /// <summary>Unordered list marker used when publishing a block.</summary>
    [Parameter(ParameterSetName = "Content"), ValidateSet("-", "*", "+")]
    public string? UnorderedListMarker { get; set; }
    string? IMarkdownWriteOptionSource.LineEnding => LineEnding?.ToString();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        MarkdownSaveOptions options = MarkdownOptionUtilities.BuildWriteOptions(this) ?? new MarkdownSaveOptions();
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowLoss))) options.AllowLoss = AllowLoss.IsPresent;
        RunDocument(Path, isNew: true, Force.IsPresent,
            path => path == null ? MarkdownDocument.Create() : MarkdownDocument.Create(path),
            (document, path, _) => { document.Save(path, options); return null; });
    }
    /// <inheritdoc />
    protected override void InvokeContent(MarkdownDocument document) {
        using (MarkdownDslContext.Enter(document)) Content!.Invoke();
    }
}
