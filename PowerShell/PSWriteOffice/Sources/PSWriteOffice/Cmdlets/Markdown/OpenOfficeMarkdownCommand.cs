using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Opens native Markdown from a file or parses text; a file content block saves atomically.</summary>
/// <example><summary>Open a literal path from a file pipeline.</summary>
/// <code>$document = Get-Item -LiteralPath './Review[final].md' | Open-OfficeMarkdown</code></example>
/// <example><summary>Parse text with bounded reader options.</summary>
/// <code>Open-OfficeMarkdown -Text '# Report' -Options (New-OfficeMarkdownReaderOptions -MaxInputCharacters 500000)</code></example>
[Cmdlet(VerbsCommon.Open, "OfficeMarkdown", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(MarkdownDocument), typeof(FileInfo))]
public sealed class OpenOfficeMarkdownCommand : OfficeLifecycleCmdlet<MarkdownDocument> {
    /// <summary>Markdown source with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;
    /// <summary>Literal source path, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;
    /// <summary>Markdown to parse without a file or authoring block.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Native reader options followed by optional native-property hashtable overrides.</summary>
    [Parameter, ValidateNotNull]
    public object[]? Options { get; set; }
    /// <summary>Native settings used when saving a file content block.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath")]
    [OfficeOptionsTransformation(typeof(MarkdownSaveOptions)), ValidateNotNull]
    public MarkdownSaveOptions? WriteOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        MarkdownLoadOptions? reader = MarkdownOptionUtilities.BindReaderOptions(engine, Options);
        if (ParameterSetName == "Text") {
            if (Content != null || MyInvocation.BoundParameters.ContainsKey(nameof(PassThru)) ||
                MyInvocation.BoundParameters.ContainsKey(nameof(Show)) || MyInvocation.BoundParameters.ContainsKey(nameof(AllowLoss))) {
                InvalidArgument(new PSArgumentException("Text only parses a document; content blocks and save controls require a file path."), Text);
                return;
            }
            WriteObject(MarkdownDocument.Parse(Text, reader));
            return;
        }
        if (Content == null && (WriteOptions != null || MyInvocation.BoundParameters.ContainsKey(nameof(AllowLoss)))) {
            InvalidArgument(new PSArgumentException("Save settings require a content block; use Save-OfficeMarkdown for a live document."), Path);
            return;
        }
        MarkdownSaveOptions writer = WriteOptions?.Clone() ?? new MarkdownSaveOptions();
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowLoss))) writer.AllowLoss = AllowLoss.IsPresent;
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string source in OfficePath.ResolveInput(SessionState, literal ? LiteralPath : Path, literal)) {
            RunDocument(source, isNew: false, force: false, path => MarkdownDocument.Load(path!, reader),
                (document, path, _) => { document.Save(path, writer); return null; });
        }
    }
    /// <inheritdoc />
    protected override void InvokeContent(MarkdownDocument document) {
        using (MarkdownDslContext.Enter(document)) Content!.Invoke();
    }
}
